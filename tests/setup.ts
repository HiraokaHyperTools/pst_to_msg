// Jest のセットアップファイル
// グローバルなテスト設定やモックをここに記述

// 必要に応じてグローバルなbeforeAll, afterAll等を設定
beforeAll(() => {
  // テスト開始前の共通処理
});

afterAll(() => {
  // テスト終了後の共通処理
});

// カスタムマッチャーの追加例
expect.extend({
  toBeWithinRange(received: number, floor: number, ceiling: number) {
    const pass = received >= floor && received <= ceiling;
    if (pass) {
      return {
        message: () =>
          `expected ${received} not to be within range ${floor} - ${ceiling}`,
        pass: true,
      };
    } else {
      return {
        message: () =>
          `expected ${received} to be within range ${floor} - ${ceiling}`,
        pass: false,
      };
    }
  },
});
