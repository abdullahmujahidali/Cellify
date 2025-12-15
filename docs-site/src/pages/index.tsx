import type {ReactNode} from 'react';
import clsx from 'clsx';
import Link from '@docusaurus/Link';
import useDocusaurusContext from '@docusaurus/useDocusaurusContext';
import Layout from '@theme/Layout';
import HomepageFeatures from '@site/src/components/HomepageFeatures';
import Heading from '@theme/Heading';
import CodeBlock from '@theme/CodeBlock';

import styles from './index.module.css';

function HomepageHeader() {
  const {siteConfig} = useDocusaurusContext();
  return (
    <header className={clsx('hero hero--primary', styles.heroBanner)}>
      <div className="container">
        <Heading as="h1" className="hero__title">
          {siteConfig.title}
        </Heading>
        <p className="hero__subtitle">{siteConfig.tagline}</p>
        <div className={styles.buttons}>
          <Link
            className="button button--secondary button--lg"
            to="/docs">
            Get Started
          </Link>
          <Link
            className="button button--outline button--lg button--secondary"
            to="/Cellify/demo/">
            Try Live Demo
          </Link>
        </div>
      </div>
    </header>
  );
}

function QuickStartSection() {
  const installCode = `npm install cellify`;
  const exampleCode = `import { Workbook, workbookToXlsxBlob } from 'cellify';

// Create a workbook
const workbook = new Workbook();
const sheet = workbook.addSheet('Sales');

// Add headers
sheet.cell(0, 0).value = 'Product';
sheet.cell(0, 1).value = 'Revenue';

// Style headers
sheet.applyStyle('A1:B1', {
  font: { bold: true, color: '#FFFFFF' },
  fill: { type: 'pattern', pattern: 'solid', foregroundColor: '#059669' }
});

// Add data
sheet.cell(1, 0).value = 'Widget';
sheet.cell(1, 1).value = 15000;

// Export to Excel
const blob = workbookToXlsxBlob(workbook);`;

  return (
    <section className={styles.quickStart}>
      <div className="container">
        <div className="row">
          <div className="col col--6">
            <Heading as="h2">Quick Start</Heading>
            <p>Install Cellify with npm, yarn, or pnpm:</p>
            <CodeBlock language="bash">{installCode}</CodeBlock>
            <p style={{marginTop: '1rem'}}>
              Cellify works in both Node.js and browsers. Create Excel files with just a few lines of code.
            </p>
          </div>
          <div className="col col--6">
            <CodeBlock language="typescript" title="Example">{exampleCode}</CodeBlock>
          </div>
        </div>
      </div>
    </section>
  );
}

export default function Home(): ReactNode {
  const {siteConfig} = useDocusaurusContext();
  return (
    <Layout
      title="Lightweight Excel Library"
      description="A lightweight, zero-dependency spreadsheet library for JavaScript and TypeScript. Create, read, and manipulate Excel files with ease.">
      <HomepageHeader />
      <main>
        <HomepageFeatures />
        <QuickStartSection />
      </main>
    </Layout>
  );
}
