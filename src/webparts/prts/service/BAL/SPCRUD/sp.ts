// // spConfig.ts
// import { spfi, SPFx } from "@pnp/sp";
// import "@pnp/sp/webs";
// import "@pnp/sp/lists";
// import "@pnp/sp/items";
// import "@pnp/sp/files";
// import { WebPartContext } from "@microsoft/sp-webpart-base";

// let _sp: ReturnType<typeof spfi> | null = null;

// export const getSP = (context: WebPartContext) => {
//   if (!_sp) {
//     _sp = spfi().using(SPFx(context));
//   }
//   return _sp;
// };

// export const setupSP = (baseUrl: string): void => {
//   _sp = spfi(baseUrl);
// };