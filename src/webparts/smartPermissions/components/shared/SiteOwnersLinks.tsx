import * as React from 'react';
import { Link } from '@fluentui/react-components';

export interface SiteOwnersLinksProps {
  owners: { title: string; email: string }[];
}

// "Site Owners: Alice, Bob." fragment with mailto links, appended to
// permission-denied message bars. Renders nothing when the list is empty.
export const SiteOwnersLinks: React.FC<SiteOwnersLinksProps> = ({ owners }) => {
  if (owners.length === 0) return null;
  return (
    <>
      {' '}Site Owners: {owners.map((o, i) => (
        <React.Fragment key={o.email || o.title}>
          {i > 0 && ', '}
          {o.email
            ? <Link href={`mailto:${o.email}`}>{o.title}</Link>
            : o.title}
        </React.Fragment>
      ))}.
    </>
  );
};
