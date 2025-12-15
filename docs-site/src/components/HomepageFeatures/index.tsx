import type {ReactNode} from 'react';
import clsx from 'clsx';
import Heading from '@theme/Heading';
import styles from './styles.module.css';

type FeatureItem = {
  title: string;
  icon: string;
  description: ReactNode;
};

const FeatureList: FeatureItem[] = [
  {
    title: 'Zero Dependencies',
    icon: '📦',
    description: (
      <>
        Only ~8KB gzipped. Cellify uses just <a href="https://github.com/101arrowz/fflate">fflate</a> for
        ZIP compression. No bloated dependencies, no security vulnerabilities from transitive packages.
      </>
    ),
  },
  {
    title: 'Full TypeScript Support',
    icon: '🔷',
    description: (
      <>
        Written in TypeScript with complete type definitions. Get full IntelliSense,
        autocomplete, and compile-time type checking in your IDE.
      </>
    ),
  },
  {
    title: 'Excel Import & Export',
    icon: '📊',
    description: (
      <>
        Read and write <code>.xlsx</code> files with full support for styling, formulas,
        merged cells, freeze panes, and more. Round-trip your Excel files.
      </>
    ),
  },
  {
    title: 'Rich Styling',
    icon: '🎨',
    description: (
      <>
        Apply fonts, colors, borders, fills, alignment, and number formats.
        Create professional-looking spreadsheets programmatically.
      </>
    ),
  },
  {
    title: 'CSV Support',
    icon: '📝',
    description: (
      <>
        Import and export CSV files with automatic delimiter detection,
        type inference for numbers and dates, and RFC 4180 compliance.
      </>
    ),
  },
  {
    title: 'Works Everywhere',
    icon: '🌐',
    description: (
      <>
        Use in Node.js for server-side generation or in browsers for client-side
        Excel creation. Same API, same results.
      </>
    ),
  },
];

function Feature({title, icon, description}: FeatureItem) {
  return (
    <div className={clsx('col col--4')}>
      <div className="text--center padding-horiz--md">
        <div className={styles.featureIcon}>{icon}</div>
        <Heading as="h3">{title}</Heading>
        <p>{description}</p>
      </div>
    </div>
  );
}

export default function HomepageFeatures(): ReactNode {
  return (
    <section className={styles.features}>
      <div className="container">
        <div className="row">
          {FeatureList.map((props, idx) => (
            <Feature key={idx} {...props} />
          ))}
        </div>
      </div>
    </section>
  );
}
