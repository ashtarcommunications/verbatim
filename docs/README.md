# docs.paperlessdebate.com

This is the documentation website for Verbatim and associated tools, hosted at docs.paperlessdebate.com

This website is built using [Docusaurus 2](https://docusaurus.io/), a modern static website generator.

### Installation

```
$ yarn
```

### Local Development

```
$ yarn start
```

This command starts a local development server and opens up a browser window. Most changes are reflected live without having to restart the server.

### Build

```
$ yarn build
```

This command generates static content into the `build` directory and can be served using any static contents hosting service.

### Deployment

Using SSH:

```
$ USE_SSH=true yarn deploy
```

Not using SSH:

```
$ GIT_USER=<Your GitHub username> yarn deploy
```

If yarn isn't globally installed or you get SSL errors:

```
NODE_OPTION=--openssl-legacy-provider GIT_USER=<Your GitHub username> USE_SSH=false npx yarn deploy
```

If you are using GitHub pages for hosting, this command is a convenient way to build the website and push to the `gh-pages` branch.
