import * as React from 'react';
import styles from './Profile.module.scss';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { MSGraphClientV3 } from '@microsoft/sp-http';
import { GraphService, Me, Person } from '../services/GraphService';

export interface ProfileProps { context: WebPartContext; }
export default function Profile(props: ProfileProps) {
  const [me, setMe] = React.useState<Me | null>(null);
  const [people, setPeople] = React.useState<Person[]>([]);
  const [error, setError] = React.useState<string | null>(null);
  const [loading, setLoading] = React.useState(true);

  React.useEffect(() => {
    let mounted = true;
    (async () => {
      try {
        const client: MSGraphClientV3 = await props.context.msGraphClientFactory.getClient('3');
        const svc = new GraphService(client);
        const [m, p] = await Promise.all([
          svc.getMe(),
          svc.getPeopleTopN(12).catch((e) => {
            // Permission fallback – still show me, suppress org list
            console.warn('People fetch failed (likely permissions):', e);
            return [];
          })
        ]);
        if (!mounted) return;
        setMe(m); setPeople(p);
      } catch (e: any) {
        if (mounted) setError(e?.message ?? 'Unknown error');
      } finally {
        if (mounted) setLoading(false);
      }
    })();
    return () => { mounted = false; };
  }, []);

  if (loading) return <div>Loading profile…</div>;
  if (error)   return <div style={{color:'#a80000'}}>Error: {error}</div>;
  if (!me)     return <div>No profile data.</div>;

  const renderSkills = (skills: {displayName:string; proficiency?:string}[]) => {
    if (!skills?.length) return <span style={{color:'#777'}}>No skills listed</span>;
    return (
      <div className={styles['card--skills']}>
        {skills.slice(0, 12).map((s, i) =>
          <span key={i} className={styles.skill}>
            {s.displayName}{s.proficiency ? ` • ${s.proficiency}` : ''}
          </span>
        )}
      </div>
    );
  };

  const Card = (p: Person | Me, top = false) => (
    <li key={p.id} className={styles.card}>
      <div className={styles['card--image']}>
        <img src={p.photoUrl ?? `https://static2.sharepointonline.com/files/fabric/office-ui-fabric-core/9.6.1/images/persona/size72.png`} alt={p.displayName}/>
      </div>
      <div className={styles['card--name']}>{p.displayName}</div>
      <div className={styles['card--meta']}>
        {p.jobTitle ?? ''}{p.jobTitle && p.department ? ' • ' : ''}{p.department ?? ''}
      </div>

      {top && (me as Me).aboutMe && (
        <div style={{marginBottom:'8px', color:'#333'}}>{(me as Me).aboutMe}</div>
      )}

      {top && (me as Me).responsibilities?.length ? (
        <div style={{marginBottom:'8px'}}>
          <strong>Ask me about:</strong>{' '}
          {(me as Me).responsibilities!.slice(0,6).join(', ')}
        </div>
      ) : null}

      {renderSkills(p.skills)}

      {'mail' in p && p.mail && (
        <div className={styles['card--email']}><a href={`mailto:${p.mail}`}>{p.mail}</a></div>
      )}
    </li>
  );

  return (
    <>
      {/* Me — hero card */}
      <ul className={styles['template--cards']} style={{background:'#fff', gridTemplateColumns:'1fr'}}>
        {Card(me, true)}
      </ul>

      {/* Org people */}
         <div className={styles.peopleScroll}>
        <ul className={styles['template--cards']}>
          {people.map(p => Card(p))}
        </ul>
      </div>
    </>
  );
}
