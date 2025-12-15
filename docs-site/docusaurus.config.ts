import {themes as prismThemes} from 'prism-react-renderer';
import type {Config} from '@docusaurus/types';
import type * as Preset from '@docusaurus/preset-classic';

const config: Config = {
  title: 'Cellify',
  tagline: 'A lightweight, zero-dependency spreadsheet library for JavaScript and TypeScript',
  favicon: 'img/favicon.ico',

  future: {
    v4: true,
  },

  // GitHub Pages deployment
  url: 'https://abdullahmujahidali.github.io',
  baseUrl: '/Cellify/',

  organizationName: 'abdullahmujahidali',
  projectName: 'Cellify',
  deploymentBranch: 'gh-pages',
  trailingSlash: false,

  onBrokenLinks: 'throw',

  markdown: {
    onBrokenMarkdownLinks: 'warn',
  },

  i18n: {
    defaultLocale: 'en',
    locales: ['en'],
  },

  presets: [
    [
      'classic',
      {
        docs: {
          sidebarPath: './sidebars.ts',
          editUrl: 'https://github.com/abdullahmujahidali/Cellify/tree/main/docs-site/',
          routeBasePath: 'docs',
        },
        blog: false, // Disable blog
        theme: {
          customCss: './src/css/custom.css',
        },
      } satisfies Preset.Options,
    ],
  ],

  themeConfig: {
    image: 'img/cellify-social-card.png',
    colorMode: {
      defaultMode: 'light',
      respectPrefersColorScheme: true,
    },
    navbar: {
      title: 'Cellify',
      logo: {
        alt: 'Cellify Logo',
        src: 'img/logo.svg',
      },
      items: [
        {
          type: 'docSidebar',
          sidebarId: 'docsSidebar',
          position: 'left',
          label: 'Documentation',
        },
        {
          href: 'https://abdullahmujahidali.github.io/Cellify/demo/',
          label: 'Live Demo',
          position: 'left',
        },
        {
          href: 'https://github.com/abdullahmujahidali/Cellify',
          label: 'GitHub',
          position: 'right',
        },
        {
          href: 'https://www.npmjs.com/package/cellify',
          label: 'npm',
          position: 'right',
        },
      ],
    },
    footer: {
      style: 'dark',
      links: [
        {
          title: 'Documentation',
          items: [
            {
              label: 'Getting Started',
              to: '/docs/getting-started',
            },
            {
              label: 'API Reference',
              to: '/docs/api/workbook',
            },
          ],
        },
        {
          title: 'Guides',
          items: [
            {
              label: 'Excel Import/Export',
              to: '/docs/guides/excel',
            },
            {
              label: 'Styling Cells',
              to: '/docs/guides/styling',
            },
          ],
        },
        {
          title: 'More',
          items: [
            {
              label: 'Live Demo',
              href: 'https://abdullahmujahidali.github.io/Cellify/demo/',
            },
            {
              label: 'GitHub',
              href: 'https://github.com/abdullahmujahidali/Cellify',
            },
            {
              label: 'npm',
              href: 'https://www.npmjs.com/package/cellify',
            },
          ],
        },
      ],
      copyright: `Copyright © ${new Date().getFullYear()} Cellify. Built with Docusaurus.`,
    },
    prism: {
      theme: prismThemes.github,
      darkTheme: prismThemes.dracula,
      additionalLanguages: ['typescript', 'bash'],
    },
  } satisfies Preset.ThemeConfig,
};

export default config;
