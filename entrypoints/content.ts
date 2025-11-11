const modifierPasteSkip = new WeakMap<Document, boolean>();

export default defineContentScript({
  matches: ["*://mail.google.com/*", "*://calendar.google.com/*"],
  allFrames: true,
  runAt: "document_idle",
  main() {
    setupDocument(document);
  },
});

function setupDocument(doc: Document) {
  doc.addEventListener("keydown", handleKeyDown, true);
  doc.addEventListener(
    "paste",
    (e) => handlePaste(e as ClipboardEvent, doc),
    true
  );
}

function handleKeyDown(event: KeyboardEvent) {
  if (!isPasteKey(event) || !event.shiftKey) return;

  const doc =
    (event.target instanceof Element ? event.target.ownerDocument : document) ??
    document;
  modifierPasteSkip.set(doc, true);
  setTimeout(() => modifierPasteSkip.delete(doc), 400);
}

function isPasteKey(event: KeyboardEvent): boolean {
  return (event.metaKey || event.ctrlKey) && event.key?.toLowerCase() === "v";
}

function handlePaste(event: ClipboardEvent, doc: Document) {
  if (!event.clipboardData) return;

  const selection = doc.getSelection();
  if (!selection || !hasValidSelection(selection)) return;

  const url = extractUrl(event.clipboardData);
  if (!url) return;

  if (modifierPasteSkip.get(doc)) {
    modifierPasteSkip.delete(doc);
    return;
  }

  const range = selection.getRangeAt(0).cloneRange();
  const selectedText = range.toString();
  if (!selectedText) return;

  event.preventDefault();
  event.stopPropagation();

  restoreSelection(selection, range);

  const existingAnchor = findEnclosingAnchor(range);
  if (existingAnchor && rangeCoversNode(range, existingAnchor)) {
    existingAnchor.href = url;
    selectNode(selection, existingAnchor);
    return;
  }

  createLink(range, url, selectedText, doc, selection);
}

function hasValidSelection(selection: Selection): boolean {
  return selection.rangeCount > 0 && !selection.isCollapsed;
}

function extractUrl(clipboardData: DataTransfer): string | null {
  const text = clipboardData.getData("text/plain").trim();
  if (!text) return null;

  try {
    const url = new URL(text);
    return url.protocol === "http:" || url.protocol === "https:"
      ? url.href
      : null;
  } catch {
    return null;
  }
}

function restoreSelection(selection: Selection, range: Range) {
  selection.removeAllRanges();
  selection.addRange(range);
}

function findEnclosingAnchor(range: Range): HTMLAnchorElement | null {
  const startAnchor = findClosestAnchor(range.startContainer);
  const endAnchor = findClosestAnchor(range.endContainer);
  return startAnchor && startAnchor === endAnchor ? startAnchor : null;
}

function findClosestAnchor(node: Node | null): HTMLAnchorElement | null {
  let current: Node | null = node;
  while (current) {
    if (current instanceof HTMLAnchorElement) return current;
    current = current.parentNode;
  }
  return null;
}

function rangeCoversNode(range: Range, node: Node): boolean {
  const nodeRange = node.ownerDocument?.createRange();
  if (!nodeRange) return false;

  nodeRange.selectNodeContents(node);
  return (
    range.compareBoundaryPoints(Range.START_TO_START, nodeRange) <= 0 &&
    range.compareBoundaryPoints(Range.END_TO_END, nodeRange) >= 0
  );
}

function createLink(
  range: Range,
  url: string,
  selectedText: string,
  doc: Document,
  selection: Selection
) {
  restoreSelection(selection, range);

  if (tryCreateLinkWithExecCommand(doc, url, selectedText, selection)) {
    return;
  }

  createLinkManually(range, url, selectedText, doc, selection);
}

function tryCreateLinkWithExecCommand(
  doc: Document,
  url: string,
  selectedText: string,
  selection: Selection
): boolean {
  if (!doc.execCommand("createLink", false, url)) {
    return false;
  }

  const link = findCreatedLink(selection, doc);
  if (!link) return false;

  link.textContent = selectedText;
  selectNode(selection, link);
  return true;
}

function findCreatedLink(
  selection: Selection,
  doc: Document
): HTMLAnchorElement | null {
  if (selection.rangeCount === 0) return null;

  const range = selection.getRangeAt(0);
  const container = range.commonAncestorContainer;

  if (container.nodeType === Node.TEXT_NODE) {
    return container.parentElement as HTMLAnchorElement | null;
  }

  if (container instanceof HTMLAnchorElement) {
    return container;
  }

  const walker = doc.createTreeWalker(container, NodeFilter.SHOW_ELEMENT, {
    acceptNode: (node) =>
      node instanceof HTMLAnchorElement
        ? NodeFilter.FILTER_ACCEPT
        : NodeFilter.FILTER_SKIP,
  });

  return walker.nextNode() as HTMLAnchorElement | null;
}

function createLinkManually(
  range: Range,
  url: string,
  selectedText: string,
  doc: Document,
  selection: Selection
) {
  restoreSelection(selection, range);

  const link = doc.createElement("a");
  link.href = url;
  link.textContent = selectedText;

  const fragment = range.extractContents();
  if (fragment.childNodes.length) {
    link.innerHTML = "";
    link.appendChild(fragment);
    if (!link.textContent.includes(selectedText)) {
      link.textContent = selectedText;
    }
  }

  range.insertNode(link);
  selectNode(selection, link);
}

function selectNode(selection: Selection, node: Node) {
  const range = node.ownerDocument?.createRange();
  if (!range) return;

  selection.removeAllRanges();
  range.selectNodeContents(node);
  selection.addRange(range);
}
