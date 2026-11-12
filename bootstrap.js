const PLUGIN_NAME = "Word 引文顺序";
const FTL_FILE = "word-citation-tags.ftl";
const SNAPSHOT_PREF = "extensions.word-citation-tags.snapshot";
const WORD_NS = new Set([
    "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "http://purl.oclc.org/ooxml/wordprocessingml/main",
]);
const MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const POLL_INTERVAL = 5000;
let menuID;
let columnID;
let pollTimer;
let childProcess;
let pendingSnapshot;
let snapshotScriptURL;
const linkedWordCommands = new Map();
const wordPaths = new WeakMap();
const documents = new Map();
const latestByCollection = new Map();
let latestDocument;
let originalGetApplication;
let linkedGetApplication;
let readVersion = 0;
let manualInProgress = false;
let stopped = false;
let lastAutomaticSignature;

function startup({ id, rootURI }) {
    snapshotScriptURL = rootURI + "word-snapshot.ps1";
    restoreOrder();
    for (const window of Zotero.getMainWindows()) {
        onMainWindowLoad({ window });
    }
    columnID = Zotero.ItemTreeManager.registerColumn({
        dataKey: "word-citation-order",
        label: Zotero.locale.startsWith("zh") ? "引文顺序" : "Citation Order",
        pluginID: id,
        enabledTreeIDs: ["main"],
        width: "90",
        dataProvider: item => {
            if (item.libraryID !== Zotero.Libraries.userLibraryID) return "";
            const rank = displayedDocument(Zotero.getActiveZoteroPane())?.ranks.get(item.key);
            return rank === undefined ? "" : String(rank).padStart(10, "0");
        },
        renderCell(_index, data, column, _isFirstColumn, document) {
            const cell = document.createElement("span");
            cell.className = "cell " + column.className;
            cell.textContent = data ? String(Number(data)) : "";
            return cell;
        },
    });
    menuID = Zotero.MenuManager.registerMenu({
        menuID: "word-citation-tags",
        pluginID: id,
        target: "main/menubar/tools",
        menus: [
            {
                menuType: "menuitem",
                l10nID: "word-citation-tags-open",
                onCommand: event => readAndSave(event.target.ownerGlobal),
            },
            {
                menuType: "menuitem",
                l10nID: "word-citation-tags-switch",
                onCommand: event => switchDocument(event.target.ownerGlobal)
                    .catch(error => Zotero.logError(error)),
            },
        ],
    });
    linkWordCommands();
    Zotero.uiReadyPromise.then(async () => {
        if (stopped) return;
        await associateLegacyOrder();
        if (Zotero.isWin) scheduleWordRead(0);
    }).catch(error => Zotero.logError(error));
}

function onMainWindowLoad({ window }) {
    window.MozXULElement.insertFTLIfNeeded(FTL_FILE);
}

function onMainWindowUnload({ window }) {
    window.document.querySelector('link[href="' + FTL_FILE + '"]')?.remove();
}

function shutdown() {
    stopped = true;
    pollTimer?.cancel();
    childProcess?.kill().catch(error => Zotero.logError(error));
    if (Zotero.Integration.getApplication === linkedGetApplication) {
        Zotero.Integration.getApplication = originalGetApplication;
    }
    const integration = Zotero.Integration.Interface.prototype;
    for (const [name, { original, wrapper }] of linkedWordCommands) {
        if (integration[name] === wrapper) integration[name] = original;
    }
    linkedWordCommands.clear();
    if (menuID) Zotero.MenuManager.unregisterMenu(menuID);
    if (columnID) Zotero.ItemTreeManager.unregisterColumn(columnID);
    for (const window of Zotero.getMainWindows()) {
        onMainWindowUnload({ window });
    }
}

function linkWordCommands() {
    const integration = Zotero.Integration;
    originalGetApplication = integration.getApplication;
    linkedGetApplication = function (agent, command, path) {
        const app = originalGetApplication.apply(this, arguments);
        if (agent === "WinWord" && path) wordPaths.set(app, path);
        return app;
    };
    integration.getApplication = linkedGetApplication;
    const methods = integration.Interface.prototype;
    for (const name of ["addEditCitation", "addCitation", "refresh"]) {
        const original = methods[name];
        const wrapper = async function (...args) {
            const result = await original.apply(this, args);
            if (!stopped && this._app.processorName === "Word") {
                try {
                    await queueConfirmedOrder(this, name);
                }
                catch (error) {
                    Zotero.logError(error);
                }
            }
            return result;
        };
        methods[name] = wrapper;
        linkedWordCommands.set(name, { original, wrapper });
    }
}

async function queueConfirmedOrder(command, name) {
    const path = wordPaths.get(command._app);
    if (!path) return;
    const collection = selectedCollection(Zotero.getActiveZoteroPane());
    const version = ++readVersion;
    const finished = Zotero.Integration.currentCommandPromise;
    const fromSession = name === "refresh" || !command._session.data.prefs.delayCitationUpdates;
    let citations;
    if (fromSession) {
        citations = Object.values(command._session.citationsByIndex)
            .flatMap(citation => citation.citationItems);
    }
    else {
        // Delayed citation updates can leave session indices stale after inserting fields.
        const fields = await command._doc.getFields(command._session.data.prefs.fieldType);
        citations = [];
        for (const nativeField of fields) {
            const field = await Zotero.Integration.Field.loadExisting(nativeField);
            if (field instanceof Zotero.Integration.CitationField) {
                const citation = await field.unserialize();
                if (citation?.citationItems) citations.push(...citation.citationItems);
            }
        }
    }
    // Updating Zotero's list after Word's cleanup keeps the citation command independent of the UI.
    Promise.resolve(finished).then(async () => {
        if (stopped || version !== readVersion) return;
        const ordered = await resolveCitations(citations, fromSession);
        if (stopped || version !== readVersion) return;
        lastAutomaticSignature = undefined;
        await saveOrder(path, ordered, collection);
    }).catch(error => Zotero.logError(error));
}

function restoreOrder() {
    try {
        const value = Zotero.Prefs.get(SNAPSHOT_PREF, true);
        if (!value) return;
        const saved = JSON.parse(value);
        if (Array.isArray(saved.documents)) {
            for (const [key, document] of saved.documents) {
                documents.set(key, { ...document, ranks: new Map(document.ranks) });
            }
        }
        else if (Array.isArray(saved.ranks) && saved.ranks.length) {
            documents.set(documentKey(saved.path) || "legacy", {
                path: saved.path || "",
                collectionKey: null,
                ranks: new Map(saved.ranks),
                legacy: true,
            });
        }
        indexCollections();
    }
    catch (error) {
        Zotero.logError(error);
    }
}

function documentKey(path) {
    if (!path) return "";
    if (!Zotero.isWin) return path;
    if (/^https?:\/\//i.test(path)) return path.replaceAll("\\", "/");
    return path.replaceAll("/", "\\").toLowerCase();
}

function selectedCollection(pane) {
    const collections = pane?.getSelectedCollections?.() ?? [];
    return collections.length === 1 ? collections[0] : null;
}

function displayedDocument(pane) {
    const collections = pane?.getSelectedCollections?.() ?? [];
    if (collections.length > 1) return undefined;
    if (collections.length === 1) {
        return collections[0].libraryID === Zotero.Libraries.userLibraryID
            ? latestByCollection.get(collections[0].key)
            : undefined;
    }
    return latestDocument;
}

function indexCollections() {
    latestByCollection.clear();
    latestDocument = undefined;
    for (const document of documents.values()) {
        if (document.collectionKey) latestByCollection.set(document.collectionKey, document);
        latestDocument = document;
    }
}

function inferCollection(items, preferred) {
    const candidates = new Map();
    for (const item of items) {
        const memberships = new Set();
        for (const id of item.getCollections()) {
            let collection = Zotero.Collections.get(id);
            while (collection && !memberships.has(collection.id)) {
                memberships.add(collection.id);
                collection = collection.parentID ? Zotero.Collections.get(collection.parentID) : null;
            }
        }
        for (const id of memberships) candidates.set(id, (candidates.get(id) ?? 0) + 1);
    }
    if (!candidates.size) return null;
    const maximum = Math.max(...candidates.values());
    const best = [...candidates.keys()].filter(id => candidates.get(id) === maximum);
    // Prefer the specific project when a parent and its subfolder contain the same citations.
    const ancestors = new Set();
    for (const id of best) {
        let parent = Zotero.Collections.get(id)?.parentID;
        while (parent) {
            ancestors.add(parent);
            parent = Zotero.Collections.get(parent)?.parentID;
        }
    }
    const projects = best.filter(id => !ancestors.has(id));
    if (projects.length === 1) return Zotero.Collections.get(projects[0])?.key ?? null;
    return preferred && projects.includes(preferred.id) ? preferred.key : null;
}

async function associateLegacyOrder() {
    const legacy = [...documents.values()].find(document => document.legacy);
    if (!legacy) return;
    const items = await Promise.all([...legacy.ranks.keys()].map(key =>
        Zotero.Items.getByLibraryAndKeyAsync(Zotero.Libraries.userLibraryID, key)));
    if (stopped || !legacy.legacy) return;
    legacy.collectionKey = inferCollection(items.filter(item => item && !item.deleted),
        selectedCollection(Zotero.getActiveZoteroPane()));
    delete legacy.legacy;
    await persistOrders();
}

async function chooseDocument(window) {
    const { FilePicker } = ChromeUtils.importESModule(
        "chrome://zotero/content/modules/filePicker.mjs"
    );
    const picker = new FilePicker();
    picker.init(window, "选择 Word 文件", picker.modeOpen);
    picker.appendFilter("Word 文档", "*.docx");
    if (await picker.show() !== picker.returnOK) return null;
    return { path: picker.file, citations: readDocx(window, picker.file) };
}

function readSnapshotDocument(window, snapshot) {
    if (snapshot.state === "live") {
        return {
            path: snapshot.path,
            citations: readParts(name => parseXML(window, snapshot.parts[name], name)),
        };
    }
    if (snapshot.state === "saved") {
        return { path: snapshot.path, citations: readDocx(window, snapshot.path) };
    }
    return null;
}

function scheduleWordRead(delay = POLL_INTERVAL) {
    if (stopped) return;
    pollTimer = Cc["@mozilla.org/timer;1"].createInstance(Ci.nsITimer);
    pollTimer.initWithCallback(() => {
        readActiveWord().catch(error => Zotero.debug(PLUGIN_NAME + "：" + error))
            .finally(() => scheduleWordRead());
    }, delay, Ci.nsITimer.TYPE_ONE_SHOT);
}

async function readActiveWord() {
    const window = Zotero.getMainWindow();
    if (!window || stopped || manualInProgress || Zotero.Integration.currentDoc) return;
    const version = readVersion;
    const snapshot = await readWordSnapshot(window);
    if (stopped || manualInProgress || Zotero.Integration.currentDoc || version !== readVersion) return;
    if (snapshot.state === "closed") {
        lastAutomaticSignature = undefined;
        return;
    }
    const document = readSnapshotDocument(window, snapshot);
    if (!document) return;
    const previous = documents.get(documentKey(document.path));
    const collection = selectedCollection(window.ZoteroPane);
    const signature = JSON.stringify([
        document.path,
        previous?.collectionKey ?? collection?.key,
        document.citations.map(({ id, uris, uri }) => [id, uris, uri]),
    ]);
    if (signature === lastAutomaticSignature) return;

    const ordered = await resolveCitations(document.citations);
    if (stopped || manualInProgress || Zotero.Integration.currentDoc || version !== readVersion) return;
    if (ordered.length && !ordered.some(entry => entry.item)) return;
    if (ordered.length || previous) {
        await saveOrder(document.path, ordered, collection);
    }
    lastAutomaticSignature = signature;
}

async function readAndSave(window) {
    if (manualInProgress || stopped) return;
    manualInProgress = true;
    try {
        const collection = selectedCollection(window.ZoteroPane);
        const document = await chooseDocument(window);
        if (!document || stopped) return;
        const version = ++readVersion;
        const ordered = await resolveCitations(document.citations);
        if (stopped || version !== readVersion) return;

        const matched = ordered.filter(entry => entry.item).length;
        if (ordered.length && !matched) {
            throw new Error("这份稿件的引文未匹配到「我的文库」中的条目。上次的引文顺序仍可查看。");
        }
        if (ordered.length || documents.has(documentKey(document.path))) {
            await saveOrder(document.path, ordered, collection);
        }

        const unmatched = ordered.filter(entry => !entry.item);
        const details = unmatched.slice(0, 8).map(entry => entry.number + " " + entry.title);
        const message = ordered.length
            ? "已保存 " + matched + "/" + ordered.length + " 条引文的顺序。"
            : "这份文件中没有找到活动的 Zotero 引文字段。";
        Services.prompt.alert(window, PLUGIN_NAME, details.length
            ? message + "\n\n未匹配的引文\n" + details.join("\n")
            : message);
    }
    catch (error) {
        Zotero.logError(error);
        if (!stopped) Services.prompt.alert(window, PLUGIN_NAME, error.message || String(error));
    }
    finally {
        manualInProgress = false;
    }
}

function readWordSnapshot(window) {
    if (!pendingSnapshot) {
        pendingSnapshot = queryWord(window).finally(() => { pendingSnapshot = undefined; });
    }
    return pendingSnapshot;
}

async function queryWord(window) {
    const { Subprocess } = ChromeUtils.importESModule("resource://gre/modules/Subprocess.sys.mjs");
    const windows = Services.dirsvc.get("WinD", Ci.nsIFile).path;
    const command = PathUtils.join(windows, "System32", "WindowsPowerShell", "v1.0", "powershell.exe");
    const script = await Zotero.File.getResourceAsync(snapshotScriptURL);
    const process = await Subprocess.call({
        command,
        arguments: ["-NoLogo", "-NoProfile", "-NonInteractive", "-Command", script],
    });
    childProcess = process;
    if (stopped) {
        await process.kill();
        return { state: "busy" };
    }
    const timeout = window.setTimeout(() => {
        process.kill(0).catch(error => Zotero.logError(error));
    }, 12000);
    try {
        let output = "";
        for (let chunk; (chunk = await process.stdout.readString());) {
            output += chunk;
        }
        const { exitCode } = await process.wait();
        return exitCode === 0 && output.trim()
            ? JSON.parse(output)
            : { state: "busy" };
    }
    finally {
        window.clearTimeout(timeout);
        childProcess = undefined;
    }
}

async function saveOrder(path, ordered, preferredCollection) {
    const key = documentKey(path);
    if (!key) return;
    const previous = documents.get(key);
    const matched = ordered.filter(entry => entry.item);
    if (ordered.length && !matched.length) return;
    const ranks = new Map(matched.map(entry => [entry.item.key, entry.number]));
    const collectionKey = previous?.collectionKey
        ?? inferCollection(matched.map(entry => entry.item), preferredCollection);
    if (previous && previous === latestDocument && previous.collectionKey === collectionKey
        && previous.ranks.size === ranks.size
        && [...ranks].every(([itemKey, rank]) => previous.ranks.get(itemKey) === rank)) return;

    // Reinsert to remember the most recently read manuscript in each collection.
    documents.delete(key);
    documents.set(key, { path, collectionKey, ranks });
    await persistOrders();
}

async function switchDocument(window) {
    const collection = selectedCollection(window.ZoteroPane);
    const entries = [...documents].reverse().filter(([, document]) =>
        !collection || !document.collectionKey || document.collectionKey === collection.key);
    if (!entries.length) {
        Services.prompt.alert(window, PLUGIN_NAME, "当前分类还没有稿件的引文顺序。");
        return;
    }
    const names = entries.map(([, document]) => document.path.split(/[\\/]/).at(-1) || "上次读取的稿件");
    const labels = names.map((name, index) => names.indexOf(name) === names.lastIndexOf(name)
        ? name : entries[index][1].path);
    const selection = { value: 0 };
    const message = collection ? "选择要在此分类显示的稿件" : "选择要显示的稿件";
    if (!Services.prompt.select(window, PLUGIN_NAME, message, labels, selection)) return;
    const [key, document] = entries[selection.value];
    if (!document.collectionKey && collection?.libraryID === Zotero.Libraries.userLibraryID) {
        document.collectionKey = collection.key;
    }
    documents.delete(key);
    documents.set(key, document);
    await persistOrders();
}

async function persistOrders() {
    indexCollections();
    const saved = [...documents].map(([key, document]) => [
        key, { ...document, ranks: [...document.ranks] },
    ]);
    Zotero.Prefs.set(SNAPSHOT_PREF, JSON.stringify({ documents: saved }), true);
    for (const window of Zotero.getMainWindows()) {
        const view = window.ZoteroPane?.itemsView;
        if (!view?.tree) continue;
        view.invalidateRowCache(true);
        if (view.getSortField() === columnID) {
            await view.refresh({ forceSortAll: true, restoreSelection: true, restoreScroll: true });
        }
        else {
            view.tree.invalidate();
        }
    }
}

function parseXML(window, text, name) {
    if (text === undefined) return null;
    const xml = new window.DOMParser().parseFromString(text, "application/xml");
    if (xml.documentElement.localName === "parsererror") {
        throw new Error("无法解析 Word 文件部件 " + name + "。");
    }
    return xml.documentElement;
}

function readParts(readXML) {
    const root = readXML("document");
    const body = Array.from(root?.children ?? []).find(node => isWord(node, "body"));
    if (!body) throw new Error("Word 文件中找不到正文。");
    const notes = {};
    for (const kind of ["footnote", "endnote"]) {
        const part = readXML(kind + "s");
        notes[kind] = new Map(Array.from(part?.children ?? [])
            .filter(node => isWord(node, kind))
            .map(node => [wordAttribute(node, "id"), node]));
    }
    return extractCitations(body, notes);
}

function readDocx(window, path) {
    if (!/\.docx$/i.test(path)) throw new Error("请选择 .docx 格式的 Word 文件。");
    const zip = Cc["@mozilla.org/libjar/zip-reader;1"].createInstance(Ci.nsIZipReader);
    const readXML = name => {
        const entry = "word/" + name + ".xml";
        if (!zip.hasEntry(entry)) return null;
        const input = Cc["@mozilla.org/binaryinputstream;1"].createInstance(Ci.nsIBinaryInputStream);
        input.setInputStream(zip.getInputStream(entry));
        try {
            const bytes = Uint8Array.from(input.readByteArray(input.available()));
            return parseXML(window, new TextDecoder().decode(bytes), name);
        }
        finally {
            input.close();
        }
    };

    zip.open(Zotero.File.pathToFile(path));
    try {
        return readParts(readXML);
    }
    finally {
        zip.close();
    }
}

function isWord(node, name) {
    return WORD_NS.has(node.namespaceURI) && node.localName === name;
}

function wordAttribute(node, name) {
    return node.getAttributeNS(node.namespaceURI, name);
}

function extractCitations(body, notes) {
    const citations = [];
    const seenNotes = new Set();

    function addField(instruction) {
        const marker = /^\s*(?:ADDIN\s+)?ZOTERO_ITEM\s+CSL_CITATION\s*/i.exec(instruction);
        if (!marker) return;
        try {
            const payload = JSON.parse(instruction.slice(marker[0].length));
            if (!Array.isArray(payload.citationItems) || !payload.citationItems.length
                || payload.citationItems.some(item => !item || typeof item !== "object")) {
                throw new Error("引文字段缺少条目");
            }
            citations.push(...payload.citationItems);
        }
        catch {
            throw new Error("发现无法解析的 Zotero 引文字段，已停止更新该稿件。");
        }
    }

    function scan(root) {
        const fields = [];
        function visit(node) {
            if (node.namespaceURI === MC_NS && node.localName === "AlternateContent") {
                const children = Array.from(node.children);
                const branch = children.find(child => child.localName === "Choice")
                    ?? children.find(child => child.localName === "Fallback");
                if (branch) visit(branch);
                return;
            }
            if (WORD_NS.has(node.namespaceURI)) {
                switch (node.localName) {
                    case "del":
                    case "moveFrom":
                        return;
                    case "fldSimple":
                        addField(wordAttribute(node, "instr") ?? "");
                        break;
                    case "fldChar": {
                        const type = wordAttribute(node, "fldCharType");
                        const field = fields.at(-1);
                        if (type === "begin") fields.push({ instruction: "", parsed: false });
                        if (type === "separate" && field && !field.parsed) {
                            addField(field.instruction);
                            field.parsed = true;
                        }
                        if (type === "end" && field) {
                            if (!field.parsed) addField(field.instruction);
                            fields.pop();
                        }
                        break;
                    }
                    case "instrText": {
                        const field = fields.at(-1);
                        if (field && !field.parsed) field.instruction += node.textContent;
                        break;
                    }
                    case "footnoteReference":
                    case "endnoteReference": {
                        const kind = node.localName === "footnoteReference" ? "footnote" : "endnote";
                        const id = wordAttribute(node, "id");
                        const key = `${kind}/${id}`;
                        const note = notes[kind].get(id);
                        if (note && !seenNotes.has(key)) {
                            seenNotes.add(key);
                            scan(note);
                        }
                        break;
                    }
                }
            }
            for (const child of node.children) visit(child);
        }
        visit(root);
        if (fields.some(field => /ZOTERO_ITEM/i.test(field.instruction))) {
            throw new Error("发现未闭合的 Zotero 引文字段，已停止更新该稿件。");
        }
    }

    scan(body);
    return citations;
}

async function resolveCitations(citations, fromWordSession = false) {
    const libraryID = Zotero.Libraries.userLibraryID;
    const cached = new Map();
    const matchedAliases = new Map();
    const references = [];
    for (const citation of citations) {
        const rawURIs = citation.uris ?? citation.uri;
        const uris = (Array.isArray(rawURIs) ? rawURIs : [rawURIs])
            .filter(uri => typeof uri === "string"
                && /^https?:\/\/zotero\.org\/(?:users\/(?:\d+|local\/\w+)|groups\/\d+)\/items\/[A-Z0-9]{8}\/?$/i.test(uri))
            .map(uri => uri.replace(/^https?:/i, "http:").replace(/\/$/, ""));
        let item;
        if (fromWordSession && Number.isSafeInteger(citation.id)) {
            const resolved = await Zotero.Items.getAsync(citation.id);
            if (resolved && resolved.libraryID === libraryID && !resolved.deleted && resolved.isRegularItem()) {
                item = resolved;
            }
        }
        for (const uri of uris) {
            if (item) break;
            if (!cached.has(uri)) {
                const resolved = await Zotero.URI.getURIItem(uri);
                cached.set(uri, resolved && resolved.libraryID === libraryID
                    && !resolved.deleted && resolved.isRegularItem() ? resolved : null);
            }
            if (cached.get(uri)) {
                item = cached.get(uri);
                break;
            }
        }
        if (item) {
            for (const uri of uris) matchedAliases.set(uri, item);
        }
        references.push({ citation, uris, item });
    }

    const ordered = [];
    const seen = new Set();
    for (const reference of references) {
        const { citation, uris } = reference;
        const item = reference.item ?? uris.map(uri => matchedAliases.get(uri)).find(Boolean);
        const keys = [...uris];
        if (item) keys.push(`local-item:${item.key}`);
        if (!keys.length && citation.id != null) keys.push(`document-item:${citation.id}`);
        const alreadySeen = keys.some(key => seen.has(key));
        for (const key of keys) seen.add(key);
        if (!alreadySeen) {
            ordered.push({
                item,
                number: ordered.length + 1,
                title: String(citation.itemData?.title || "题名未知"),
            });
        }
    }
    return ordered;
}
