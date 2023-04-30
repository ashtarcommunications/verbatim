// @ts-check
// Note: type annotations allow type checking and IDEs autocompletion

const lightCodeTheme = require('prism-react-renderer/themes/github');
const darkCodeTheme = require('prism-react-renderer/themes/dracula');

/** @type {import('@docusaurus/types').Config} */
const config = {
  title: 'Paperless Debate Manual',
  tagline: 'Verbatim. openCaselist. More.',
  favicon: 'img/favicon.ico',

  // Set the production url of your site here
  url: 'https://docs.paperlessdebate.com',
  // Set the /<baseUrl>/ pathname under which your site is served
  // For GitHub pages deployment, it is often '/<projectName>/'
  baseUrl: '/',
  trailingSlash: false,

  // GitHub pages deployment config.
  // If you aren't using GitHub pages, you don't need these.
  organizationName: 'ashtarcommunications', // Usually your GitHub org/user name.
  projectName: 'verbatim', // Usually your repo name.

  onBrokenLinks: 'throw',
  onBrokenMarkdownLinks: 'warn',

  scripts: [{src: 'https://analytics.aaronhardy.net/js/script.file-downloads.js', defer: true, 'data-domain': 'docs.paperlessdebate.com'}],

  // Even if you don't use internalization, you can use this field to set useful
  // metadata like html lang. For example, if your site is Chinese, you may want
  // to replace "en" with "zh-Hans".
  i18n: {
    defaultLocale: 'en',
    locales: ['en'],
  },

  presets: [
    [
      'classic',
      /** @type {import('@docusaurus/preset-classic').Options} */
      ({
        docs: {
          routeBasePath: '/',
          sidebarPath: require.resolve('./sidebars.js'),
          sidebarCollapsed: false,
          // Please change this to your repo.
          // Remove this to remove the "edit this page" links.
          // editUrl:
          //   'https://github.com/ashtarcommunications/verbatim',
        },
        theme: {
          customCss: require.resolve('./src/css/custom.css'),
        },
      }),
    ],
  ],

  themeConfig:
    /** @type {import('@docusaurus/preset-classic').ThemeConfig} */
    ({
      image: 'img/verbatim.ico',
      navbar: {
        title: 'Paperless Debate Manual',
        logo: {
          alt: 'Paperless Debate Logo',
          src: 'img/verbatim.ico',
        },
        items: [
          {
            href: 'https://paperlessdebate.com',
            position: 'right',
            'aria-label': 'Back to paperlessdebate.com',
            label: 'Back to paperlessdebate.com',
          },
        ],
      },
      footer: {
        style: 'dark',
        links: [
          {
            title: 'Links',
            items: [
              {
                label: 'GitHub',
                href: 'https://github.com/ashtarcommunications/verbatim',
              },
              {
                label: 'Donate',
                href: 'https://paperlessdebate.com/donate',
              },
            ],
          },
        ],
        copyright: `Copyright © ${new Date().getFullYear()} Ashtar Communications`,
      },
      prism: {
        theme: lightCodeTheme,
        darkTheme: darkCodeTheme,
      },
    }),
};

module.exports = config;
