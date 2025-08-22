import * as React from 'react';
import styles from './Profile.module.scss';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { MSGraphClientV3 } from '@microsoft/sp-http';
import { GraphService, Me, Person, Skill } from '../services/GraphService';

export interface ProfileProps { context: WebPartContext; }

export default function Profile(props: ProfileProps) {
  const [me, setMe] = React.useState<Me | null>(null);
  const [people, setPeople] = React.useState<Person[]>([]);
  const [next, setNext] = React.useState<string | undefined>(undefined);
  const [error, setError] = React.useState<string | null>(null);
  const [loading, setLoading] = React.useState(true);
  const [loadingMore, setLoadingMore] = React.useState(false);

  // ==== Skill sorting helpers (EN + DE, punctuation-safe) ====

  type Rule = { rank: number; rx: RegExp };

  // Order matters: highest rank first
  const LEVEL_RULES: Rule[] = [
    { rank: 5, rx: /\b(expert|experte|expertin|principal|architect)\b/i },
    { rank: 4, rx: /\b(advanced|fortgeschritten|senior|professional|profi|specialist)\b/i },
    { rank: 3, rx: /\b(associate|intermediate|mittel|mittelstufe)\b/i },
    { rank: 2, rx: /\b(foundation|fundamentals|basic|grundkenntnisse)\b/i },
    { rank: 1, rx: /\b(junior|beginner|einsteiger|newbie)\b/i },
  ];

  // Extra: catch patterns like " - Expert", "— Expert", ": Expert"
  const TRAILING_EXPERT = /[\s:\-–—]\s*expert\b/i;
  const TRAILING_ADV    = /[\s:\-–—]\s*(advanced|fortgeschritten)\b/i;

  const emailFor = (p: Person | Me) => encodeURIComponent(p.mail || p.userPrincipalName);

  const outlookNewMeeting = (p: Person | Me) =>
    `https://outlook.office.com/calendar/deeplink/compose?to=${emailFor(p)}&subject=${encodeURIComponent('Termin mit ' + p.displayName)}`;

  const teamsChat = (p: Person | Me) =>
    `https://teams.microsoft.com/l/chat/0/0?users=${emailFor(p)}`;

  // move to property pane later if you want:
  const CONSULTANT_PROFILES_URL =
    'https://thinformatics.sharepoint.com/sites/Beraterprofile/Freigegebene%20Dokumente/Forms/AllItems.aspx?as=json';


  function textRank(t?: string): number {
    if (!t) return 2; // neutral
    // quick path for common trailing forms
    if (TRAILING_EXPERT.test(t)) return 5;
    if (TRAILING_ADV.test(t)) return 4;
    for (const rule of LEVEL_RULES) {
      if (rule.rx.test(t)) return rule.rank;
    }
    return 2;
  }

  function rankForSkill(s: Skill): number {
    // Prefer structured proficiency when it clearly maps
    const pRank = textRank(s.proficiency);
    if (pRank !== 2) return pRank;
    // Fall back to the display text (e.g., "Exchange Online - Expert")
    return textRank(s.displayName);
  }

 function sortSkillsByLevel(skills: Skill[]): Skill[] {
    return [...skills].sort((a, b) => {
      const rb = rankForSkill(b);
      const ra = rankForSkill(a);
      if (rb !== ra) return rb - ra; // higher rank first
      // stable-ish tiebreaker
      return (a.displayName || '').localeCompare(b.displayName || '', undefined, { sensitivity: 'base' });
    });
  }

  // Optional: show a derived label when proficiency is missing
 function effectiveProficiency(s: Skill): string | undefined {
    if (s.proficiency) return s.proficiency;
    const r = rankForSkill(s);
    if (r >= 3) {
      // pretty labels for derived levels
      if (r === 5) return 'Expert';
      if (r === 4) return 'Advanced';
      if (r === 3) return 'Associate';
    }
    return undefined;
  }

  // modal state
  const [skillsModal, setSkillsModal] = React.useState<{ name: string; skills: Skill[] } | null>(null);

  const scrollRef = React.useRef<HTMLDivElement>(null);

  const loadFirstPage = React.useCallback(async () => {
    try {
      const client: MSGraphClientV3 = await props.context.msGraphClientFactory.getClient('3');
      const svc = new GraphService(client);

      const meData = await svc.getMe();
      setMe(meData);

      // Try directory (active Members only). If it fails (no consent), fallback to /me/people.
      try {
        const { items, nextLink } = await svc.getPeoplePage(24);
        setPeople(items);
        setNext(nextLink);
      } catch (dirErr: any) {
        console.warn('Directory read failed, using /me/people:', dirErr);
        const { items } = await svc.getPeopleFallback(24);
        setPeople(items);
        setNext(undefined);
      }
    } catch (e: any) {
      setError(e?.message ?? 'Unknown error');
    } finally {
      setLoading(false);
    }
  }, [props.context.msGraphClientFactory]);

  const loadMore = React.useCallback(async () => {
    if (!next || loadingMore) return;
    setLoadingMore(true);
    try {
      const client: MSGraphClientV3 = await props.context.msGraphClientFactory.getClient('3');
      const svc = new GraphService(client);
      const { items, nextLink } = await svc.getPeoplePage(24, next);
      setPeople(prev => [...prev, ...items]);
      setNext(nextLink);
    } catch (e: any) {
      console.warn('Load more failed:', e);
    } finally {
      setLoadingMore(false);
    }
  }, [next, loadingMore, props.context.msGraphClientFactory]);

  React.useEffect(() => { loadFirstPage(); }, [loadFirstPage]);

  const onScroll = React.useCallback((e: React.UIEvent<HTMLDivElement>) => {
    if (!next || loadingMore) return;
    const el = e.currentTarget;
    const nearBottom = el.scrollTop + el.clientHeight >= el.scrollHeight - 200;
    if (nearBottom) loadMore();
  }, [next, loadingMore, loadMore]);

  // close modal on ESC
  React.useEffect(() => {
    const onKey = (ev: KeyboardEvent) => { if (ev.key === 'Escape') setSkillsModal(null); };
    window.addEventListener('keydown', onKey);
    return () => window.removeEventListener('keydown', onKey);
  }, []);

  if (loading) return <div>Loading profile…</div>;
  if (error)   return <div style={{ color: '#a80000' }}>Error: {error}</div>;
  if (!me)     return <div>No profile data.</div>;

  const renderSkillsCompact = (p: Person | Me) => {
    const skills = p.skills ?? [];
    if (!skills.length) return <span style={{ color: '#777' }}>No skills listed</span>;

    const sorted = sortSkillsByLevel(skills);            // sort by strength
    const visible = skills.slice(0, 12);
    const hiddenCount = Math.max(0, skills.length - visible.length);

    return (
      <>
        <div className={styles['card--skills']}>
          {visible.map((s, i) =>
            <span key={i} className={styles.skill}>
              {s.displayName}
             {effectiveProficiency(s) ? ` • ${effectiveProficiency(s)}` : ''}  {/* optional label */}
            </span>
          )}
        </div>
        {hiddenCount > 0 && (
          <button
            className={styles.showAllBtn}
            onClick={() => setSkillsModal({ name: p.displayName, skills: sorted })}  // pass sorted
            aria-label={`Show all ${skills.length} skills for ${p.displayName}`}
          >
            Alle anzeigen ({sorted.length})
          </button>
        )}
      </>
    );
  };
  
  const Card = (p: Person | Me, top: boolean = false) => {
  // NOTE: return a single <li> with a matching </li>
  return (
    <li key={p.id} className={styles.card}>
      <div className={styles['card--image']}>
        <img
          src={p.photoUrl ?? 'https://static2.sharepointonline.com/files/fabric/office-ui-fabric-core/9.6.1/images/persona/size72.png'}
          alt={p.displayName}
        />
      </div>
  
      <div className={styles['card--name']}>{p.displayName}</div>
  
      <div className={styles['card--meta']}>
        {p.jobTitle ?? ''}{p.jobTitle && p.department ? ' • ' : ''}{p.department ?? ''}
      </div>
  
      {top && (me as Me)?.aboutMe && (
        <div style={{ marginBottom: '8px', color: '#333' }}>{(me as Me).aboutMe}</div>
      )}
  
      {top && (me as Me)?.responsibilities?.length ? (
        <div style={{ marginBottom: '8px' }}>
          <strong>Ask me about:</strong>{' '}
          {(me as Me).responsibilities!.slice(0, 6).join(', ')}
        </div>
      ) : null}

      {/* email (always) */}
      <div className={styles['card--email']}>
        <a href={`mailto:${p.mail || p.userPrincipalName}`}>
          {p.mail || p.userPrincipalName}
        </a>
      </div>
  
      {/* actions */}
      <div className={styles['card--links']}>
        <a
          className={styles.linkBtn}
          href={outlookNewMeeting(p)}
          target="_blank"
          rel="noopener noreferrer"
          title={`Termin mit ${p.displayName}`}
        >
          <img src="https://thinformatics.sharepoint.com/:i:/r/sites/thinformationHub/SiteAssets/SitePages/Skill-Search/32px-Microsoft_Office_Outlook_(2018%E2%80%93present).svg.png?csf=1&web=1&e=AVZl0q" alt="Outlook Logo" className={styles.logo} />

           Termin
        </a>
  
        <a
          className={styles.linkBtn}
          href={teamsChat(p)}
          target="_blank"
          rel="noopener noreferrer"
          title={`Teams-Chat mit ${p.displayName}`}
        >
          <img src="https://thinformatics.sharepoint.com/:i:/r/sites/thinformationHub/SiteAssets/SitePages/Skill-Search/32px-Microsoft_Office_Teams_(2018%E2%80%93present).svg.png?csf=1&web=1&e=bABdsE" alt="Word Logo" className={styles.logo} /> 

           Chat
        </a>
  
        {/* Stage 1: show to everyone; Stage 2 will gate this */}
        <a
          className={styles.linkBtn}
          href={CONSULTANT_PROFILES_URL}
          target="_blank"
          rel="noopener noreferrer"
          title="Berater-Profil"
        >
          <img src="https://thinformatics.sharepoint.com/:i:/r/sites/thinformationHub/SiteAssets/SitePages/Skill-Search/32px-Microsoft_Office_SharePoint_(2019%E2%80%93present).svg.png?csf=1&web=1&e=etkaPW" alt="SharePoint Logo" className={styles.logo} />
            Profil anzeigen
        </a>

      </div>
      {/* skills */}
      {renderSkillsCompact(p)}
    </li>
  );
  };
  
  return (
    <>
      {/* Me — hero card */}
      <ul className={styles['template--cards']} style={{ background: '#fff', gridTemplateColumns: '1fr' }}>
        {Card(me, true)}
      </ul>

      {/* Org people (scrollable) */}
      <div ref={scrollRef} className={styles.peopleScroll} onScroll={onScroll}>
        <ul className={styles['template--cards']}>
          {people.map(p => Card(p))}
          {loadingMore && <li>Loading more…</li>}
        </ul>
      </div>

      {/* Skills modal */}
      {skillsModal && (
        <div className={styles.modalBackdrop} onClick={() => setSkillsModal(null)}>
          <div className={styles.modalCard} onClick={(e) => e.stopPropagation()}>
            <div className={styles.modalHeader}>
              Skills — {skillsModal.name}
              <button className={styles.modalClose} onClick={() => setSkillsModal(null)} aria-label="Close">×</button>
            </div>
            <div className={styles.modalBody}>
              <div className={`${styles['card--skills']} ${styles['card--skills--full']}`}>
                 {sortSkillsByLevel(skillsModal.skills).map((s, i) =>
                  <span key={i} className={styles.skill}>
                    {s.displayName}
                    {effectiveProficiency(s) ? ` • ${effectiveProficiency(s)}` : ''}
                  </span>
                )}
              </div>
            </div>
          </div>
        </div>
      )}
    </>
  );
}
