import { WebPartContext } from "@microsoft/sp-webpart-base";
import { spfi, SPFx, SPFI } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/attachments";
import "@pnp/sp/fields";
import { IToDoItem, ILookupOption, IAttachment } from "../models/IToDoItem";

export class SPService {
    private _sp: SPFI;
    private _fieldMap: Map<string, string> = new Map();
    // ─── NEW: stores TypeAsString for each internal name ──────────────────────
    private _fieldTypeMap: Map<string, string> = new Map();
    private _listTitle: string = "";
    private _isInitialized: boolean = false;
    private _lastError: string = "";

    constructor(context: WebPartContext) {
        this._sp = spfi().using(SPFx(context));
    }

    public async init(): Promise<void> {
        if (this._isInitialized) return;
        try {
            const urlName = "ToDo";
            const lists = await this._sp.web.lists.select("Title", "RootFolder/Name").expand("RootFolder")();
            const match = lists.find(l =>
                l.Title.toLowerCase() === urlName.toLowerCase() ||
                l.RootFolder.Name.toLowerCase() === urlName.toLowerCase() ||
                l.Title.replace(/\s+/g, '').toLowerCase() === urlName.toLowerCase()
            );
            this._listTitle = match ? match.Title : "ToDo";

            console.log(`[SPService] List Identified: "${this._listTitle}" (searching for "${urlName}")`);

            const fields = await this._sp.web.lists.getByTitle(this._listTitle).fields
                .select("Title", "InternalName", "TypeAsString")();

            // Store BOTH display→internal name map AND internal→type map with trimmed keys
            fields.forEach(f => {
                const title = (f.Title || "").trim();
                const internal = (f.InternalName || "").trim();
                const type = (f.TypeAsString || "").trim();

                this._fieldMap.set(title, internal);
                this._fieldTypeMap.set(internal, type);
            });

            console.log("[SPService] Discovered Fields Mapping:", Array.from(this._fieldMap.entries()));
            console.log("[SPService] Field Types Mapping:", Array.from(this._fieldTypeMap.entries()));

            this._isInitialized = true;
        } catch (error) {
            console.error("[SPService] Initialization Failed:", error);
            this._listTitle = "ToDo";
            this._isInitialized = true;
        }
    }

    public getDiscoveredFields(): any {
        const obj: any = { ListTitle: this._listTitle };
        this._fieldMap.forEach((v, k) => { obj[k] = v; });
        return obj;
    }

    public getLastError(): string { return this._lastError; }

    public getInternalName(title: string, fallback: string): string {
        if (this._fieldMap.has(title)) return this._fieldMap.get(title)!;
        const keys = Array.from(this._fieldMap.keys());
        const match = keys.find(k =>
            k.toLowerCase() === title.toLowerCase() ||
            k.replace(/\s/g, '').toLowerCase() === title.toLowerCase()
        );
        if (match) return this._fieldMap.get(match)!;
        return fallback;
    }

    public async checkMissingFields(): Promise<{ title: string; internalName: string; exists: boolean }[]> {
        await this.init();
        const required = [
            { title: "Subject", internal: "Subject" },
            { title: "Category", internal: "Category" },
            { title: "Status", internal: "Status" },
            { title: "Priority", internal: "Priority" },
            { title: "Due Date", internal: "DueDate" },
            { title: "Task Owner", internal: "TaskOwner" },
        ];
        const fieldInternalNames = Array.from(this._fieldMap.values());
        const fieldTitles = Array.from(this._fieldMap.keys());
        return required.map(r => ({
            title: r.title,
            internalName: r.internal,
            exists: fieldTitles.some(t => t.toLowerCase() === r.title.toLowerCase()) ||
                    fieldInternalNames.some(i => i.toLowerCase() === r.internal.toLowerCase())
        }));
    }

    public async getToDoItems(): Promise<IToDoItem[]> {
        await this.init();

        const fieldInternalNames = Array.from(this._fieldMap.values());

        const names = {
            Subject:          this.getInternalName("Subject",          "Title"),
            TaskOwner:        this.getInternalName("Task Owner",       "TaskOwner"),
            AssigneeInternal: this.getInternalName("Assigne Internal", "AssigneInternal"),
            AssigneeExternal: this.getInternalName("Assigne External", "AssigneExternal"),
            Status:           this.getInternalName("Status",           "Status"),
            Category:         this.getInternalName("Category",         "Category"),
            Classification:   this.getInternalName("Classification",   "Classification"),
            Priority:         this.getInternalName("Priority",         "Priority"),
            CompletedPercent: this.getInternalName("Completed %",      "CompletedPercent"),
            StartDate:        this.getInternalName("Start Date",       "StartDate"),
            CompletionDate:   this.getInternalName("Completion Date",  "CompletionDate"),
            CreatedByUser:    this.getInternalName("Created By User",  "CreatedByUser"),
            UpdatedByUser:    this.getInternalName("Updated By User",  "UpdatedByUser"),
            CreatedOn:        this.getInternalName("Created On",       "CreatedOn"),
            UpdatedOn:        this.getInternalName("Updated On",       "UpdatedOn"),
            Description:      this.getInternalName("Description",      "Description"),
            Regarding:        this.getInternalName("Regarding",        "Regarding"),
            DueDate:          this.getInternalName("Due Date",         "DueDate"),
            Resolution:       this.getInternalName("Resolution",       "Resolution"),
            EmailNotifications: this.getInternalName("Email Notification", "EmailNotifications"),
        };

        const selects = ["Id", "Title", "Author/Title", "Author/EMail", "Editor/Title", "Editor/EMail", "Created", "Modified"];
        const expands = ["Author", "Editor"];

        // ─── Use ACTUAL TypeAsString to decide expand strategy ─────────────────
        const safelyAddSelect = (internalName: string): void => {
            if (fieldInternalNames.indexOf(internalName) < 0) return;
            const fieldType = this._fieldTypeMap.get(internalName) || "";

            if (fieldType === "User" || fieldType === "UserMulti") {
                selects.push(`${internalName}/Id`, `${internalName}/Title`, `${internalName}/EMail`);
                expands.push(internalName);
            } else if (fieldType === "Lookup" || fieldType === "LookupMulti") {
                selects.push(`${internalName}/Id`, `${internalName}/Title`);
                expands.push(internalName);
            } else {
                selects.push(internalName);
            }
        };

        safelyAddSelect(names.Subject);
        safelyAddSelect(names.Status);
        safelyAddSelect(names.Category);
        safelyAddSelect(names.Classification);
        safelyAddSelect(names.Priority);
        safelyAddSelect(names.TaskOwner);
        safelyAddSelect(names.AssigneeInternal);
        safelyAddSelect(names.AssigneeExternal);
        safelyAddSelect(names.StartDate);
        safelyAddSelect(names.CompletionDate);
        safelyAddSelect(names.CompletedPercent);
        safelyAddSelect(names.CreatedByUser);
        safelyAddSelect(names.UpdatedByUser);
        safelyAddSelect(names.CreatedOn);
        safelyAddSelect(names.UpdatedOn);
        safelyAddSelect(names.Description);
        safelyAddSelect(names.Regarding);
        safelyAddSelect(names.DueDate);
        safelyAddSelect(names.Resolution);
        safelyAddSelect(names.EmailNotifications);

        try {
            const rawItems = await this._sp.web.lists
                .getByTitle(this._listTitle).items
                .select(...selects)
                .expand(...expands)();

            return rawItems.map(item => {

                // ── Handles Lookup (object) and Choice (string) ──────────────
                const getLookup = (n: string): { Id: number; Title: string } | undefined => {
                    const v = item[n];
                    if (v === null || v === undefined || v === "") return undefined;
                    if (typeof v === "object") return { Id: v.Id || 0, Title: v.Title || "" };
                    if (typeof v === "string") return { Id: 0, Title: v };
                    return undefined;
                };

                // ── Handles Person field ──────────────────────────────────────
                const getPerson = (n: string): { Id: number; Title: string; EMail: string } | undefined => {
                    const v = item[n];
                    if (!v || typeof v !== "object") return undefined;
                    return { Id: v.Id || 0, Title: v.Title || "", EMail: v.EMail || "" };
                };

                return {
                    Id: item.Id,
                    Title: item[names.Subject] || item.Title || "",
                    Description: item[names.Description],
                    Status:         getLookup(names.Status),
                    Category:       getLookup(names.Category),
                    Classification: getLookup(names.Classification),
                    Priority:       getLookup(names.Priority),
                    TaskOwner:        getPerson(names.TaskOwner),
                    AssigneeInternal: getPerson(names.AssigneeInternal),
                    AssigneeExternal: getPerson(names.AssigneeExternal),
                    Regarding:        item[names.Regarding],
                    DueDate:          item[names.DueDate],
                    StartDate:        item[names.StartDate],
                    CompletionDate:   item[names.CompletionDate],
                    CompletedPercent: item[names.CompletedPercent],
                    EmailNotifications: item[names.EmailNotifications],
                    Author:  getPerson(names.CreatedByUser)  || item.Author,
                    Editor:  getPerson(names.UpdatedByUser)  || item.Editor,
                    Created: item[names.CreatedOn]  || item.Created,
                    Modified: item[names.UpdatedOn] || item.Modified,
                    Resolution: item[names.Resolution],
                };
            });
        } catch (error) {
            console.error(`Query failed on ${this._listTitle}:`, error);
            this._lastError = `Fetch failed: ${error.message || JSON.stringify(error)}`;
            try {
                const fallback = await this._sp.web.lists
                    .getByTitle(this._listTitle).items.select("Id", "Title", "Created")();
                return fallback.map(i => ({ Id: i.Id, Title: i.Title, Created: i.Created }));
            } catch { return []; }
        }
    }

    public async addToDoItem(item: any): Promise<any> {
        await this.init();
        try {
            const fieldInternalNames = Array.from(this._fieldMap.values());
            const cleaned: any = {};

            const currentUser = await this._sp.web.currentUser();
            const now = new Date().toISOString();
            const createdOnInt    = this.getInternalName("Created On",       "CreatedOn");
            const createdByInt    = this.getInternalName("Created By User",  "CreatedByUser");
            const updatedOnInt    = this.getInternalName("Updated On",       "UpdatedOn");
            const updatedByInt    = this.getInternalName("Updated By User",  "UpdatedByUser");

            if (fieldInternalNames.indexOf(createdOnInt)  > -1) item[createdOnInt]          = now;
            if (fieldInternalNames.indexOf(createdByInt)  > -1) item[`${createdByInt}Id`]   = currentUser.Id;
            if (fieldInternalNames.indexOf(updatedOnInt)  > -1) item[updatedOnInt]           = now;
            if (fieldInternalNames.indexOf(updatedByInt)  > -1) item[`${updatedByInt}Id`]   = currentUser.Id;

            Object.keys(item).forEach(key => {
                const baseKey = key.endsWith("Id") ? key.slice(0, -2) : key;
                if (
                    fieldInternalNames.indexOf(key)     > -1 ||
                    fieldInternalNames.indexOf(baseKey) > -1 ||
                    key === "Title"
                ) {
                    if (item[key] !== undefined && item[key] !== null) cleaned[key] = item[key];
                }
            });

            return await this._sp.web.lists.getByTitle(this._listTitle).items.add(cleaned);
        } catch (error) {
            this._lastError = `Save failed: ${error.message || JSON.stringify(error)}`;
            throw error;
        }
    }

    public async updateToDoItem(id: number, item: any): Promise<any> {
        await this.init();
        try {
            const fieldInternalNames = Array.from(this._fieldMap.values());
            const cleaned: any = {};

            const currentUser  = await this._sp.web.currentUser();
            const now          = new Date().toISOString();
            const updatedOnInt = this.getInternalName("Updated On",       "UpdatedOn");
            const updatedByInt = this.getInternalName("Updated By User",  "UpdatedByUser");

            if (fieldInternalNames.indexOf(updatedOnInt) > -1) item[updatedOnInt]          = now;
            if (fieldInternalNames.indexOf(updatedByInt) > -1) item[`${updatedByInt}Id`]  = currentUser.Id;

            Object.keys(item).forEach(key => {
                const baseKey = key.endsWith("Id") ? key.slice(0, -2) : key;
                if (
                    fieldInternalNames.indexOf(key)     > -1 ||
                    fieldInternalNames.indexOf(baseKey) > -1 ||
                    key === "Title"
                ) {
                    if (item[key] !== undefined && item[key] !== null) cleaned[key] = item[key];
                }
            });

            return await this._sp.web.lists.getByTitle(this._listTitle).items.getById(id).update(cleaned);
        } catch (error) {
            if (error.data) {
                const data = await error.data.json();
                this._lastError = `Update failed: ${data.odata.error.message.value}`;
            }
            throw error;
        }
    }

    public async getLookupOptions(listUrlName: string, displayField: string = "Title"): Promise<ILookupOption[]> {
        try {
            const realTitle = await this.findListTitle(listUrlName);
            const items = await this._sp.web.lists.getByTitle(realTitle).items
                .select("Id", displayField, "Title")();
            return items.map(item => ({ Id: item.Id, Title: item[displayField] || item.Title }));
        } catch { return []; }
    }

    public async getAttachments(itemId: number): Promise<IAttachment[]> {
        await this.init();
        try { return await this._sp.web.lists.getByTitle(this._listTitle).items.getById(itemId).attachmentFiles(); }
        catch { return []; }
    }

    public async uploadAttachment(itemId: number, file: File): Promise<void> {
        await this.init();
        await this._sp.web.lists.getByTitle(this._listTitle).items.getById(itemId).attachmentFiles.add(file.name, file);
    }

    public async deleteAttachment(itemId: number, fileName: string): Promise<void> {
        await this.init();
        await this._sp.web.lists.getByTitle(this._listTitle).items.getById(itemId).attachmentFiles.getByName(fileName).delete();
    }

    public async deleteToDoItem(id: number): Promise<void> {
        await this.init();
        await this._sp.web.lists.getByTitle(this._listTitle).items.getById(id).delete();
    }

    private async findListTitle(urlName: string): Promise<string> {
        try {
            const lists = await this._sp.web.lists.select("Title", "RootFolder/Name").expand("RootFolder")();
            const match = lists.find(l =>
                l.Title.toLowerCase() === urlName.toLowerCase() ||
                l.RootFolder.Name.toLowerCase() === urlName.toLowerCase() ||
                l.Title.replace(/\s+/g, '').toLowerCase() === urlName.toLowerCase()
            );
            return match ? match.Title : urlName;
        } catch { return urlName; }
    }
}
