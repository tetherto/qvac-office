declare module 'bare-png' {
  export function encode(image: { width: number; height: number; data: Uint8Array }): Uint8Array
  export function decode(png: Uint8Array): { width: number; height: number; data: Uint8Array }
}
