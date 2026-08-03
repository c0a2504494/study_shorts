# iOS・Android版の開発

Study ShortsはGoogle Apps Scriptを実行環境として使わず、CapacitorのWebView内で動作できます。Book、学習回数、ランク、設定は端末内のIndexedDBへ保存されます。

## 必要な環境

- Node.js 20以上
- npm
- iOS: macOS、Xcode、Apple Developer設定
- Android: Android Studio、Android SDK

## 初回セットアップ

```bash
npm install
npm run build
npm run mobile:add
```

`mobile:add`は`ios/`と`android/`を生成するため、原則として初回だけ実行します。

## iOSで開く

```bash
npm run mobile:ios
```

Xcodeが開いたら、Signing & CapabilitiesでTeamとBundle Identifierを確認し、接続したiPhoneまたはSimulatorで実行します。

## Androidで開く

```bash
npm run mobile:android
```

Android Studioが開いたら、エミュレーターまたはUSB接続端末を選択して実行します。

## 静的アプリの生成

```bash
npm run build
```

次のファイルが生成されます。

- `dist/index.html`
- `dist/version.json`

ビルド処理はGASテンプレートを除去し、次のファイルを1つの静的HTMLへ統合します。

- `Index.html`
- `App.html`
- `LocalStore.html`
- `Editor.html`
- `LocalUi.html`
- `StartGuard.html`

## 保存方式

通常利用ではGoogleスプレッドシートへ接続しません。端末内のIndexedDBが正本です。

端末変更やアプリ削除に備え、アプリ内の`Export Backup`でJSONバックアップを作成してください。別端末では`Import Backup`から復元できます。

## 公開前に必要な作業

- アプリアイコンとスプラッシュ画面
- プライバシーポリシーと利用規約
- iOSの署名・Provisioning Profile
- Androidの署名鍵
- 実機テスト
- App Store Connect / Google Play Consoleへの登録
