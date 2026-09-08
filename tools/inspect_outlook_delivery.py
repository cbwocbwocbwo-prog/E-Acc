from __future__ import annotations

import win32com.client

OUTLOOK_OL_FOLDER_OUTBOX = 4
OUTLOOK_OL_FOLDER_SENT_MAIL = 5
subject_text = "[법인카드] 미등록 사용내역 등록 요청"

outlook = win32com.client.Dispatch("Outlook.Application")
namespace = outlook.GetNamespace("MAPI")
print("default_store=", namespace.DefaultStore.DisplayName)
for label, folder_id in (("outbox", OUTLOOK_OL_FOLDER_OUTBOX), ("sent", OUTLOOK_OL_FOLDER_SENT_MAIL)):
    folder = namespace.GetDefaultFolder(folder_id)
    matches = []
    for index in range(1, min(folder.Items.Count, 100) + 1):
        item = folder.Items.Item(index)
        if subject_text in str(getattr(item, "Subject", "")):
            matches.append((str(getattr(item, "SentOn", "")), str(getattr(item, "To", "")), str(getattr(item, "Subject", ""))))
    print(label, "count=", folder.Items.Count, "matching=", len(matches))
    for value in matches[:10]:
        print(value)
