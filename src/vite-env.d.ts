/// <reference types="vite/client" />
declare const __AUTOOFFICE_TEST_BRIDGE__: boolean;

declare module '*.md?raw' {
  const content: string;
  export default content;
}
