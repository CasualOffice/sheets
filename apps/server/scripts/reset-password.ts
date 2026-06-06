#!/usr/bin/env -S node --import tsx
/**
 * Operator escape hatch for Mode 3 (personal docker) — reset any
 * user's password without knowing the old one. The intended call
 * site is `docker exec`:
 *
 *   docker exec <container> casual-sheets-reset-password joel
 *
 * Reads `CASUAL_USERS_DB_PATH` (default `/data/users.db`) and writes
 * directly through `PersonalAuthStore.changePassword(userId, null,
 * newPassword)` — the `null` current-password is the documented
 * escape-hatch path on the store API.
 *
 * Modes
 *   - Interactive: no `--password` arg → prompts on stdin with echo
 *     disabled, then a confirmation prompt.
 *   - Non-interactive: `--password=<value>` (CI / scripted reset).
 *
 * Side effects:
 *   - All sessions for the target user are invalidated (every other
 *     browser/tab gets logged out on next request). This matches the
 *     web-side "change password" behaviour.
 *
 * Exit codes:
 *   0   success
 *   1   user not found
 *   2   password rejected (too short)
 *   3   DB open failed
 *   4   bad invocation (missing username)
 */
import { argv, exit, stdin, stdout } from 'node:process';
import { createInterface } from 'node:readline/promises';

import { PersonalAuthStore } from '../src/auth/personal.js';

const usage = () => {
  stdout.write(
    'Usage: casual-sheets-reset-password <username> [--password=<new-password>]\n' +
      '       casual-sheets-reset-password --help\n' +
      '\n' +
      "Resets a user's password in the Mode 3 (personal docker) auth\n" +
      'database. Reads CASUAL_USERS_DB_PATH (default /data/users.db).\n' +
      '\n' +
      'When --password is omitted the script prompts on stdin with echo\n' +
      'disabled, then asks for a confirmation. Passwords are stored as\n' +
      'bcrypt hashes; the new value never touches disk in plaintext.\n' +
      '\n' +
      'Every active session for the user is invalidated after a successful\n' +
      'reset.\n',
  );
};

async function main(): Promise<number> {
  const args = argv.slice(2);
  if (args.length === 0 || args.includes('--help') || args.includes('-h')) {
    usage();
    return args.length === 0 ? 4 : 0;
  }

  const username = args.find((a) => !a.startsWith('--'));
  if (!username) {
    usage();
    return 4;
  }

  const passwordArg = args.find((a) => a.startsWith('--password='));
  const inlinePassword = passwordArg?.slice('--password='.length);

  const dbPath = process.env.CASUAL_USERS_DB_PATH ?? '/data/users.db';

  let store: PersonalAuthStore;
  try {
    // Mode doesn't matter for the CLI — we're just reading + writing
    // existing rows. Use 'multi' so the store doesn't refuse to load
    // (it never checks mode for the changePassword path either).
    store = new PersonalAuthStore({ dbPath, mode: 'multi', bootstrap: null });
  } catch (err) {
    stdout.write(
      `error: couldn't open users.db at ${dbPath} — ${err instanceof Error ? err.message : String(err)}\n`,
    );
    return 3;
  }

  try {
    const userId = store.findIdByUsername(username);
    if (userId == null) {
      stdout.write(`error: no user named "${username}"\n`);
      return 1;
    }

    let newPassword: string;
    if (inlinePassword) {
      newPassword = inlinePassword;
    } else {
      const rl = createInterface({
        input: stdin,
        output: stdout,
        // No history file in the container.
        historySize: 0,
      });
      try {
        const first = await promptHidden(rl, `New password for ${username}: `);
        const second = await promptHidden(rl, 'Confirm new password: ');
        if (first !== second) {
          stdout.write('error: passwords do not match\n');
          return 2;
        }
        newPassword = first;
      } finally {
        rl.close();
      }
    }

    if (newPassword.length < 8) {
      stdout.write('error: password must be at least 8 characters\n');
      return 2;
    }

    const ok = store.changePassword(userId, null, newPassword);
    if (!ok) {
      stdout.write('error: password rejected (too short or user gone)\n');
      return 2;
    }

    stdout.write(
      `ok: reset password for "${username}" — every active session for this user has been invalidated.\n`,
    );
    return 0;
  } finally {
    store.close();
  }
}

/** Prompt with echo suppressed via raw-mode tricks so the password
 *  doesn't end up in shell history or visible to a shoulder-surfer. */
async function promptHidden(
  rl: ReturnType<typeof createInterface>,
  prompt: string,
): Promise<string> {
  stdout.write(prompt);
  // Hide echoed characters: replace stdout.write while reading the
  // line, so node:readline can still see backspaces.
  const isTTY = stdout.isTTY && stdin.isTTY;
  const realWrite = stdout.write.bind(stdout);
  const guard = (chunk: unknown, ...rest: unknown[]): boolean => {
    // Swallow echoes while the readline is active. The prompt itself
    // was already written; readline's per-keystroke echoes go via
    // stream.write — masking those is what hides the password.
    if (typeof chunk === 'string' && (chunk === '\r\n' || chunk === '\n' || chunk === '\r')) {
      return realWrite(chunk, ...(rest as []));
    }
    return true;
  };
  if (isTTY) {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    (stdout as any).write = guard;
  }
  try {
    const answer = await rl.question('');
    return answer;
  } finally {
    if (isTTY) {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      (stdout as any).write = realWrite;
    }
  }
}

void main().then(
  (code) => exit(code),
  (err) => {
    stdout.write(`fatal: ${err instanceof Error ? (err.stack ?? err.message) : String(err)}\n`);
    exit(3);
  },
);
