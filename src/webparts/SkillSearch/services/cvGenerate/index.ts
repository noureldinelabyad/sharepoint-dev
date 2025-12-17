import { SPHttpClient } from '@microsoft/sp-http';
//import { BERATERPROFIL_SITE, LIB_INTERNAL_NAME } from '../profileConstants';
import { BERATERPROFIL_SITE } from '../profileConstants';

import { buildDataportDocx } from './docx/dataportFallback';
import { fillDataportTemplate } from './docx/templateRenderer';
import { templateHasDocxtemplaterTags } from './docx/templateDetection';
export { downloadArrayBuffer } from './download';
export { MISSING_TOKEN } from '../constants';
export { extractProfileDataFromDocx } from './profileExtractor';
export { ProjectItem, ProfileData, SkillGroup } from './types';

export async function tryGenerateDataportDoc(
  spHttp: SPHttpClient,
  maybeTemplateUrlOrServerRel: string | null,
  data: import('./types').ProfileData
): Promise<Blob> {
  if (maybeTemplateUrlOrServerRel) {
    try {
      return await fillDataportTemplate(spHttp, maybeTemplateUrlOrServerRel, data);
    } catch (err) {
      console.warn('Falling back to generated Dataport layout because template rendering failed', err);
    }
  }
  return buildDataportDocx(data);
}

export function defaultTemplateUrl(): string {
  //return `${BERATERPROFIL_SITE}/${LIB_INTERNAL_NAME}/Dataport CV Vorlage - TAGGED.docx`;
  return `${BERATERPROFIL_SITE}/Dataport CV Vorlage - TAGGED.docx`;

}

export { buildDataportDocx, fillDataportTemplate, templateHasDocxtemplaterTags };
