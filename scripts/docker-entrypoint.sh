#!/bin/sh
# Copyright 2026 Casual Office
#
# Licensed under the Apache License, Version 2.0 (the "License");
# you may not use this file except in compliance with the License.
# You may obtain a copy of the License at
#
#     http://www.apache.org/licenses/LICENSE-2.0
#
# Unless required by applicable law or agreed to in writing, software
# distributed under the License is distributed on an "AS IS" BASIS,
# WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
# See the License for the specific language governing permissions and
# limitations under the License.

set -e

# Personal mode (Phase C) opens a SQLite users.db under the data root. When that
# root is a BIND MOUNT to a host directory owned by root — Docker creates the
# host dir root-owned if it doesn't exist — the `node` user can't create the DB
# and the server dies on boot with:
#
#   SqliteError: unable to open database file  (SQLITE_CANTOPEN)
#
# The image chowns /data to `node` at BUILD time, but a bind mount masks that
# (the host dir's ownership wins at runtime). So fix ownership HERE, at runtime,
# as root, then drop privileges to `node`. (GitHub #57.)
#
# If the operator overrode the user (`docker run --user …`), we're not root and
# can't chown — respect their choice and exec as-is; matching the mount's
# ownership is then on them.

DATA_DIR="${CASUAL_LOCAL_PATH:-/data}"

if [ "$(id -u)" = "0" ]; then
  mkdir -p "$DATA_DIR" /data
  # Best-effort: a read-only mount would make chown fail — don't crash the boot.
  chown -R node:node "$DATA_DIR" /data 2>/dev/null || true
  exec su-exec node "$@"
fi

exec "$@"
