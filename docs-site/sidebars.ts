import type {SidebarsConfig} from '@docusaurus/plugin-content-docs';

const sidebars: SidebarsConfig = {
  docsSidebar: [
    'intro',
    'getting-started',
    {
      type: 'category',
      label: 'Guides',
      items: [
        'guides/excel',
        'guides/csv',
        'guides/styling',
        'guides/merging',
        'guides/formulas',
        'guides/accessibility',
        'guides/examples',
      ],
    },
    {
      type: 'category',
      label: 'API Reference',
      items: [
        'api/workbook',
        'api/sheet',
        'api/cell',
        'api/types',
      ],
    },
  ],
};

export default sidebars;
