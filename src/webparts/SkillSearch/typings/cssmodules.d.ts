// Typings for CSS modules // This quiets all “Cannot find module *.module.scss” errors in VS Code and at build time.

declare module '*.module.scss' {
  const classes: { [key: string]: string };
  export default classes;
}
