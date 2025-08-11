// Jest のカスタムマッチャーの型定義
declare global {
  namespace jest {
    interface Matchers<R> {
      toBeWithinRange(a: number, b: number): R;
    }
  }
}

export {};
