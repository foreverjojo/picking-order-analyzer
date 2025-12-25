---
description: 更新 GitHub - 每次修正 picking-order-analyzer 程式後必須執行
---

# 更新 GitHub 流程

每次對 picking-order-analyzer 專案進行修改後，都需要同步更新到 GitHub。

## 步驟

1. 確認所有變更已儲存

2. 將變更加入暫存區
```bash
git add -A
```

3. 建立 commit（描述這次的修改內容）
// turbo
```bash
git commit -m "修改描述"
```

4. 推送到 GitHub
// turbo
```bash
git push origin main
```

## 注意事項

- 如果有 revert 或 reset 操作，可能需要使用 `git push origin main --force`
- Commit 訊息應該簡潔描述這次修改的內容
- 確保程式在本地測試正常後再推送
