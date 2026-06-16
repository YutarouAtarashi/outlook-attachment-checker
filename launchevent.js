/*
 * Outlook Smart Alerts - 添付忘れチェック検証用
 * 目的: 件名・本文に添付を示唆する語句があり、添付ファイルが0件の場合に送信前警告を表示する。
 * 注意: 本ファイルは検証用スターターです。実運用前にI情/管理者レビューを受けてください。
 */

const ATTACHMENT_KEYWORDS = [
  "添付",
  "添付します",
  "添付いたします",
  "添付ファイル",
  "別添",
  "別紙",
  "資料を送付",
  "ファイルを送付",
  "データを送付",
  "PDF",
  "Excel",
  "エクセル",
  "CSV",
  "attached",
  "attachment",
  "attach",
  "enclosed",
  "please find attached",
  "see attached"
];

const EXCLUDE_KEYWORDS = [
  "添付不要",
  "添付なし",
  "添付は不要",
  "資料は後送",
  "別途送付",
  "後ほど送付",
  "ファイルなし"
];

function containsAny(text, words) {
  const normalized = (text || "").toLowerCase();
  return words.find(w => normalized.includes(String(w).toLowerCase())) || "";
}

function getSubjectAsync() {
  return new Promise(resolve => {
    try {
      Office.context.mailbox.item.subject.getAsync(result => {
        resolve(result.status === Office.AsyncResultStatus.Succeeded ? result.value || "" : "");
      });
    } catch (e) {
      resolve("");
    }
  });
}

function getBodyTextAsync() {
  return new Promise(resolve => {
    try {
      Office.context.mailbox.item.body.getAsync(Office.CoercionType.Text, result => {
        resolve(result.status === Office.AsyncResultStatus.Succeeded ? result.value || "" : "");
      });
    } catch (e) {
      resolve("");
    }
  });
}

function getAttachmentCountAsync() {
  return new Promise(resolve => {
    try {
      if (Office.context.mailbox.item.getAttachmentsAsync) {
        Office.context.mailbox.item.getAttachmentsAsync(result => {
          if (result.status === Office.AsyncResultStatus.Succeeded && Array.isArray(result.value)) {
            resolve(result.value.length);
          } else {
            resolve(0);
          }
        });
      } else if (Array.isArray(Office.context.mailbox.item.attachments)) {
        resolve(Office.context.mailbox.item.attachments.length);
      } else {
        resolve(0);
      }
    } catch (e) {
      resolve(0);
    }
  });
}

async function onMessageSendHandler(event) {
  try {
    const [subject, body, attachmentCount] = await Promise.all([
      getSubjectAsync(),
      getBodyTextAsync(),
      getAttachmentCountAsync()
    ]);

    const checkText = `${subject}\n${body}`;
    const excludeHit = containsAny(checkText, EXCLUDE_KEYWORDS);
    const hitKeyword = containsAny(checkText, ATTACHMENT_KEYWORDS);

    if (!excludeHit && hitKeyword && attachmentCount === 0) {
      event.completed({
        allowEvent: false,
        errorMessage:
          `添付忘れの可能性があります。\n\n` +
          `件名または本文に「${hitKeyword}」が含まれていますが、添付ファイルがありません。\n` +
          `添付が必要な場合は、ファイルを追加してから再度送信してください。`
      });
      return;
    }

    event.completed({ allowEvent: true });
  } catch (error) {
    // 初期検証では業務停止を避けるため、異常時は送信を許可する。
    // 本番導入時は運用方針に応じて softBlock / block を検討する。
    event.completed({ allowEvent: true });
  }
}

Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
