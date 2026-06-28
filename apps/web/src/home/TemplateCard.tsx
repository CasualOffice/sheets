/**
 * Copyright 2026 Casual Office
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

import { Icon } from '../shell/Icon';
import { TemplateThumbnail } from './TemplateThumbnail';
import type { Template } from './registry';

export function TemplateCard({
  template,
  size = 'md',
  onPick,
}: {
  template: Template;
  size?: 'md' | 'lg';
  onPick: (t: Template) => void;
}) {
  return (
    <button
      type="button"
      className={`tpl-card tpl-card--${size}`}
      data-template-id={template.id}
      data-testid={`tpl-card-${template.id}`}
      style={{ ['--tpl-accent' as string]: template.accent }}
      onClick={() => onPick(template)}
    >
      <div className="tpl-card__preview">
        <TemplateThumbnail template={template} />
      </div>
      <div className="tpl-card__body">
        <div className="tpl-card__icon">
          <Icon name={template.icon} />
        </div>
        <div className="tpl-card__text">
          <div className="tpl-card__name">{template.name}</div>
          <div className="tpl-card__category">{template.category}</div>
        </div>
      </div>
    </button>
  );
}
