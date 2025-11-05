import * as React from 'react';
import {  Stack } from '@fluentui/react';
import { SPHttpClient } from '@microsoft/sp-http';
import { saveAs } from 'file-saver';
import styles from "../../SkillSearch.module.scss";

import {
  findLatestProfileDocx,
  resolveDataportTemplateUrl
} from '../../services/profileRepo';

import {
  downloadArrayBuffer,
  extractProfileDataFromDocx,
  tryGenerateDataportDoc
} from '../../services/cvGenerate';

type Props = {
  spHttpClient: SPHttpClient;
  absWebUrl: string;           // context.pageContext.web.absoluteUrl
  serverRelWebUrl: string;     // context.pageContext.web.serverRelativeUrl
  displayName: string;         // e.g. "Noureldin Elabyad"
};

export const GernrateCv: React.FC<Props> = ({
  spHttpClient, absWebUrl, serverRelWebUrl, displayName
}) => {
  const [busy, setBusy] = React.useState(false);

  const generateCv = React.useCallback(async () => {
    if (!displayName) return;
    setBusy(true);
    try {
      // 1) Latest Beraterprofil in the person's folder
      const repo = await findLatestProfileDocx(spHttpClient, absWebUrl, displayName);
      if (!repo) { alert('Kein Beraterprofil gefunden.'); return; }

      // 2) Parse source profile from the *.docx
      const sourceBuf = await downloadArrayBuffer(spHttpClient, repo.fileUrl);
      const data = await extractProfileDataFromDocx(sourceBuf);

      // 3) Resolve template *if present* (root or person folder); we don't require Templates folder
      const tplUrl = await resolveDataportTemplateUrl(spHttpClient, absWebUrl, repo.libRootUrl, repo.folderUrl);

      // 4) Generate filled Dataport doc (use template if it has tags; else build one)
      const outBlob = await tryGenerateDataportDoc(spHttpClient, tplUrl, data);

      // 5) Download to user's default Downloads folder
      const fileName = `Dataport CV - ${displayName} - ${new Date().toISOString().slice(0, 10)}.docx`;
      saveAs(outBlob, fileName);
    } catch (e: any) {
      console.error(e);
      alert(`Erstellung fehlgeschlagen.\n${e?.message || e}`);
    } finally {
      setBusy(false);
    }
  }, [spHttpClient, absWebUrl, displayName]);

  return (
    <Stack aria-disabled={busy ? 'true' : 'false'}>
      <a 
        href="#"
        className={styles.linkBtn}
        role="button"
        aria-disabled={busy || !displayName ? 'true' : 'false'}
        onClick={(e) => {
          if (busy || !displayName) { e.preventDefault(); return; }
          e.preventDefault(); // avoid jumping to '#'
          generateCv();
        }}
        onKeyDown={(e) => {
          if ((e.key === 'Enter' || e.key === ' ') && !(busy || !displayName)) {
            e.preventDefault();
            generateCv();
          }
        }}
      >
        <img
          src="https://thinformatics.sharepoint.com/:i:/r/sites/thinformationHub/SiteAssets/SitePages/Skill-Search/32px-Microsoft_Office_SharePoint_(2019%E2%80%93present).svg.png?csf=1&web=1&e=etkaPW"
          alt=""
          className={styles.logo}
        />      
        Dataport CV generieren
      </a>
    </Stack>
  );
};

export default GernrateCv;
