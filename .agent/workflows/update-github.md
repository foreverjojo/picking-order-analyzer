---
description: 更新 GitHub - 每次修正 picking-order-analyzer 程式後必須執行
---


5. # 更新 GitHub 流程 (符合 Ivy House Rules)

每次對 picking-order-analyzer 專案進行修改後，需遵循以下安全流程同步到 GitHub。

## 步驟

1. **[關鍵] 本地測試確認**
   - 請使用者在本地瀏覽器打開頁面測試功能。
   - 詢問：「請問測試結果是否正常？可以上傳了嗎？」
   - **只有在使用者確認後，才執行後續步驟。**

2. 將變更加入暫存區
```bash
git add -A
```

3. 建立 commit
   - **注意：** 訊息須遵循 Conventional Commits (如 `feat:`, `fix:`, `docs:`)
```bash
git commit -m "feat: [描述修改內容]"
```

4. 推送到 GitHub (觸發部署)
```bash
git push origin main
```

## 注意事項

- **嚴禁**在未經使用者測試前直接 Push。
- Push 後，提醒使用者 Cloudflare Pages 會自動更新。
