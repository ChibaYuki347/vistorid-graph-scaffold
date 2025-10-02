
/**
 * Outlook Add-in (Compose) snippet
 * - Save to get itemId
 * - Store VisitorID into CustomProperties (hidden)
 * - POST minimal linkage to backend (tenant, room, itemId)
 */
declare const Office: any;

type LinkPayload = {
  tenantId: string;
  roomUpn: string;
  ewsItemId: string;
  restItemId?: string;
  visitorId: string;
};

async function saveItem(): Promise<string> {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.item.saveAsync((result: any) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value as string); // EWS-style itemId
      } else {
        reject(result.error);
      }
    });
  });
}

async function getItemId(): Promise<string> {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.item.getItemIdAsync((res: any) => {
      if (res.status === Office.AsyncResultStatus.Succeeded) {
        resolve(res.value as string);
      } else {
        reject(res.error);
      }
    });
  });
}

async function loadCustomProperties(): Promise<any> {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.item.loadCustomPropertiesAsync((res: any) => {
      if (res.status === Office.AsyncResultStatus.Succeeded) {
        resolve(res.value);
      } else {
        reject(res.error);
      }
    });
  });
}

export async function attachVisitor(tenantId: string, roomUpn: string, visitorId: string, backendBaseUrl: string) {
  // 1) Ensure the item is saved (get itemId). New appointments send no invites on saveAsync.
  let itemId: string;
  try {
    // If it's a brand-new item, save first
    await saveItem();
    itemId = await getItemId();
  } catch (e) {
    // Try again (compose save can lag on first call)
    await new Promise(r => setTimeout(r, 1500));
    await saveItem();
    itemId = await getItemId();
  }

  // 2) Optionally also compute a REST id for legacy/compat scenarios
  let restItemId: string | undefined = undefined;
  try {
    const v = Office.MailboxEnums.RestVersion.v1_0; // Graph-compatible format
    restItemId = Office.context.mailbox.convertToRestId(itemId, v);
  } catch {}

  // 3) Store hidden CustomProperties (visitorId)
  const props = await loadCustomProperties();
  props.set("visitorId", visitorId);
  await new Promise<void>((resolve, reject) => {
    props.saveAsync((res: any) => res.status === Office.AsyncResultStatus.Succeeded ? resolve() : reject(res.error));
  });

  // 4) Notify backend (it will translate IDs, upsert DB, optionally write Graph extensions)
  const payload: LinkPayload = {
    tenantId, roomUpn,
    ewsItemId: itemId,
    restItemId,
    visitorId
  };

  // NOTE: Use your authenticated channel (SSO/OBO or cookie)
  await fetch(`${backendBaseUrl}/link`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload),
    credentials: "include"
  });
}
