/* ホーム画面に「アプリとして追加」できるようにするためだけのファイル。

   Androidのブラウザは、このファイル（サービスワーカー）が登録されていて、
   かつ fetch を受け取る作りになっていないと「アプリを追加」を出さない。

   ここでは中身を一切ためこまない（キャッシュしない）。
   ためこむと、直したところが職員の画面に出てこなくなるため。
   毎回そのまま通信させ、画面の新しさは index.html 側の版番号で管理する。 */

self.addEventListener("install", () => {
  // 前のものを待たず、すぐ新しいほうに入れ替える
  self.skipWaiting();
});

self.addEventListener("activate", (e) => {
  e.waitUntil((async () => {
    // 以前にためこんだものが残っていたら、すべて消す
    try {
      const names = await caches.keys();
      await Promise.all(names.map((n) => caches.delete(n)));
    } catch (err) { /* 消せなくても動きに支障はない */ }
    await self.clients.claim();
  })());
});

self.addEventListener("fetch", (e) => {
  // 何もためこまず、そのまま通す
  e.respondWith(fetch(e.request));
});
