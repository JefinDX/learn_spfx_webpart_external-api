// A file is required to be in the root of the /src directory by the TypeScript compiler

//need to tell TS about the css modules
declare module '*.module.scss' {
    const value: Record<string, string>;
    export default value;
}