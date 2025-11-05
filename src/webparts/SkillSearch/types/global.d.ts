// Tell TS about the browser build of mammoth
declare module 'mammoth/mammoth.browser' {
  const mammoth: any;
  export = mammoth;
}

// If you still see file-saver typing errors, you can keep this too.
// (Not strictly needed once @types/file-saver is installed)
declare module 'file-saver';
