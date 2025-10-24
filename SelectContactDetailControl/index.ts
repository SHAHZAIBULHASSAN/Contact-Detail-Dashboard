import { IInputs, IOutputs } from "./generated/ManifestTypes";
import "./CSS/ContactDetailControl.css";
import { Chart, registerables, ChartConfiguration, ChartType } from "chart.js";

Chart.register(...registerables);

interface Contact {
  id: string;
  name: string;
  email: string;
  city: string;
  job: string;
}

interface RelatedRecord {
  id: string;
  name: string;
  type: string;
  status?: string;
}

type EntityCount = Record<string, number>;

export class ContactDetailControl implements ComponentFramework.StandardControl<IInputs, IOutputs> {
  private container!: HTMLDivElement;
  private chart?: Chart;
  private chartType: ChartType = "bar";
  private chartContext: CanvasRenderingContext2D | null = null;
  private selectedContactId: string | null = null;
  private contacts: Contact[] = [];
  private context!: ComponentFramework.Context<IInputs>;
  private notifyOutputChanged!: () => void;
  private entityCounts: EntityCount = {};

  public init(
    context: ComponentFramework.Context<IInputs>,
    notifyOutputChanged: () => void,
    _state: ComponentFramework.Dictionary,
    container: HTMLDivElement
  ): void {
    this.context = context;
    this.notifyOutputChanged = notifyOutputChanged;
    this.container = container;

    this.renderLayout();
    this.registerEventHandlers();
    void this.loadContacts();
  }

  private renderLayout(): void {
    this.container.innerHTML = `
      <div class="dashboard-container">
        <div class="left-panel">
          <h2>Contacts</h2>
          <input type="text" id="searchBox" placeholder="Search contacts..." />
          <div id="contactList" class="contact-list"></div>
        </div>

        <div class="right-panel">
          <div class="chart-section">
            <div class="chart-header">
              <h3>Distribution</h3>
              <div class="chart-controls">
                <select id="chartTypeSelector">
                  <option value="bar">Bar</option>
                  <option value="pie">Pie</option>
                  <option value="doughnut">Doughnut</option>
                  <option value="polarArea">Polar Area</option>
                  <option value="line">Line</option>
                  <option value="radar">Radar</option>
                </select>
                <button id="byCity" class="active">By City</button>
                <button id="byJob">By Job</button>
              </div>
            </div>
            <div class="chart-wrapper" style="width:100%;height:260px;">
              <canvas id="contactChart"></canvas>
            </div>
          </div>

          <div class="tabs-section">
            <div class="tabs-header" id="tabsHeader">
              <button class="tab active" data-tab="contact">Contact</button>
              <button class="tab" data-tab="opportunities">Opportunities</button>
              <button class="tab" data-tab="activities">Activities</button>
              <button class="tab" data-tab="orders">Orders</button>
              <button class="tab" data-tab="quotes">Quotes</button>
              <button class="tab" data-tab="products">Products</button>
              <button class="tab" data-tab="accounts">Accounts</button>
              <button class="tab" data-tab="leads">Leads</button>
            </div>
            <div class="tab-and-counts">
              <div class="tab-content" id="tabContent">Select a contact to view details.</div>
              <div id="entityCounts" class="entity-counts">
                <h4>Related Counts</h4>
                <ul></ul>
              </div>
            </div>
          </div>
        </div>
      </div>
    `;

    const canvas = this.container.querySelector("#contactChart") as HTMLCanvasElement | null;
    this.chartContext = canvas ? canvas.getContext("2d") : null;

    if (!this.chartContext) {
      console.warn("ContactDetailControl: canvas #contactChart not found during layout render.");
    }
  }

  private registerEventHandlers(): void {
    const chartTypeSelector = this.container.querySelector("#chartTypeSelector") as HTMLSelectElement | null;
    const byCityBtn = this.container.querySelector("#byCity") as HTMLButtonElement | null;
    const byJobBtn = this.container.querySelector("#byJob") as HTMLButtonElement | null;
    const searchBox = this.container.querySelector("#searchBox") as HTMLInputElement | null;

    chartTypeSelector?.addEventListener("change", (e: Event) => {
      const value = (e.target as HTMLSelectElement).value as ChartType;
      this.chartType = value;
      this.updateChart(byCityBtn?.classList.contains("active") ? "city" : "job");
    });

    byCityBtn?.addEventListener("click", () => {
      byCityBtn.classList.add("active");
      byJobBtn?.classList.remove("active");
      this.updateChart("city");
    });

    byJobBtn?.addEventListener("click", () => {
      byJobBtn.classList.add("active");
      byCityBtn?.classList.remove("active");
      this.updateChart("job");
    });

    searchBox?.addEventListener("input", (e: Event) => {
      const term = (e.target as HTMLInputElement).value.trim().toLowerCase();
      this.renderContactList(this.contacts.filter(c => c.name.toLowerCase().includes(term)));
    });

    const tabs = this.container.querySelectorAll(".tab");
    tabs.forEach(tab => {
      tab.addEventListener("click", (ev: Event) => {
        const selectedTab = (ev.currentTarget as HTMLElement).getAttribute("data-tab") || "contact";
        this.switchTab(selectedTab);
      });
    });
  }

  private async loadContacts(): Promise<void> {
    try {
      const result = await this.context.webAPI.retrieveMultipleRecords(
        "contact",
        "?$select=fullname,emailaddress1,contactid,jobtitle,address1_city"
      );

      this.contacts = result.entities.map((c: ComponentFramework.WebApi.Entity) => ({
        id: String(c.contactid ?? ""),
        name: String(c.fullname ?? "Unnamed"),
        email: String(c.emailaddress1 ?? "N/A"),
        city: String(c.address1_city ?? "Unknown"),
        job: String(c.jobtitle ?? "N/A")
      })).filter(ct => ct.id);

      this.renderContactList(this.contacts);
      this.calculateEntityCounts();
      this.updateChart("city");
    } catch (error) {
      console.error("ContactDetailControl.loadContacts error:", error);
      const tabContent = this.container.querySelector("#tabContent") as HTMLDivElement | null;
      if (tabContent) tabContent.textContent = "Failed to load contacts. Check console for details.";
    }
  }

  private renderContactList(contacts: Contact[]): void {
    const contactList = this.container.querySelector("#contactList") as HTMLDivElement | null;
    if (!contactList) return;

    contactList.innerHTML = "";
    if (!contacts.length) {
      contactList.textContent = "No contacts found.";
      return;
    }

    contacts.forEach(contact => {
      const div = document.createElement("div");
      div.className = "contact-card";
      div.setAttribute("data-id", contact.id);
      div.textContent = `${contact.name} - ${contact.email} - ${contact.city}`;
      div.onclick = () => this.onContactSelect(contact);
      contactList.appendChild(div);
    });
  }

  private onContactSelect(contact: Contact): void {
    this.selectedContactId = contact.id;
    this.switchTab("contact");
    this.showContactDetails(contact);
  }

  private switchTab(tabName: string): void {
    const tabs = this.container.querySelectorAll(".tab");
    tabs.forEach(t => t.classList.remove("active"));
    const activeTab = this.container.querySelector(`[data-tab="${tabName}"]`) as HTMLElement | null;
    activeTab?.classList.add("active");

    const tabContent = this.container.querySelector("#tabContent") as HTMLDivElement | null;
    if (!tabContent) return;

    if (!this.selectedContactId && tabName !== "contact") {
      tabContent.textContent = "Select a contact first.";
      return;
    }

    if (tabName === "contact") {
      const selectedContact = this.contacts.find(c => c.id === this.selectedContactId);
      if (selectedContact) this.showContactDetails(selectedContact);
    } else {
      void this.loadRelatedData(tabName);
    }
  }

  private showContactDetails(contact: Contact): void {
    const tabContent = this.container.querySelector("#tabContent") as HTMLDivElement | null;
    if (!tabContent) return;

    tabContent.innerHTML = `
      <h4>${this.escapeHtml(contact.name)}</h4>
      <p><strong>Email:</strong> ${this.escapeHtml(contact.email)}</p>
      <p><strong>City:</strong> ${this.escapeHtml(contact.city)}</p>
      <p><strong>Job Title:</strong> ${this.escapeHtml(contact.job)}</p>
    `;
    this.updateChart("city");
  }

  private async loadRelatedData(entityName: string): Promise<void> {
    const tabContent = this.container.querySelector("#tabContent") as HTMLDivElement | null;
    if (!tabContent) return;
    tabContent.textContent = "Loading related records...";

    try {
      const mockRelated: RelatedRecord[] = Array.from({ length: Math.floor(Math.random() * 5) + 1 }, (_, i) => ({
        id: `${entityName}-${i + 1}`,
        name: `${entityName.charAt(0).toUpperCase() + entityName.slice(1)} Record ${i + 1}`,
        type: entityName,
        status: i % 2 === 0 ? "Active" : "Closed"
      }));

      await new Promise(resolve => setTimeout(resolve, 150));
      this.renderRelatedRecords(mockRelated, entityName);
      this.updateChartByEntity(mockRelated, entityName);
    } catch (err) {
      console.error("Error loading related data:", err);
      tabContent.textContent = "Failed to load related records.";
    }
  }

  private renderRelatedRecords(records: RelatedRecord[], entityName: string): void {
    const tabContent = this.container.querySelector("#tabContent") as HTMLDivElement | null;
    if (!tabContent) return;

    tabContent.innerHTML = `<p>Total ${this.escapeHtml(entityName)}: ${records.length}</p>`;
    const ul = document.createElement("ul");
    records.forEach(r => {
      const li = document.createElement("li");
      li.textContent = `${r.name} [${r.status ?? "N/A"}]`;
      ul.appendChild(li);
    });
    tabContent.appendChild(ul);

    this.entityCounts[entityName] = records.length;
    this.renderEntityCounts();
  }

  private renderEntityCounts(): void {
    const container = this.container.querySelector("#entityCounts ul") as HTMLUListElement | null;
    if (!container) return;
    container.innerHTML = "";
    for (const [entity, count] of Object.entries(this.entityCounts)) {
      const li = document.createElement("li");
      li.textContent = `${entity}: ${count}`;
      container.appendChild(li);
    }
  }

  private updateChart(groupBy: "city" | "job"): void {
    if (!this.chartContext) return;

    const grouped: Record<string, number> = {};
    this.contacts.forEach(c => {
      const key = groupBy === "city" ? c.city : c.job;
      grouped[key] = (grouped[key] || 0) + 1;
    });

    this.renderChart(grouped, "Contacts");
  }

  private updateChartByEntity(records: RelatedRecord[], entityName: string): void {
    if (!this.chartContext) return;

    const grouped: Record<string, number> = {};
    records.forEach(r => {
      const k = r.status ?? "Unknown";
      grouped[k] = (grouped[k] || 0) + 1;
    });

    this.renderChart(grouped, entityName);
  }

  private renderChart(grouped: Record<string, number>, label = "Count"): void {
    const labels = Object.keys(grouped);
    const data = Object.values(grouped);

    if (this.chart) {
      try {
        this.chart.destroy();
      } catch (e) {
        console.warn("Error destroying previous chart:", e);
      }
    }

    if (!this.chartContext) return;

    const config: ChartConfiguration = {
      type: this.chartType,
      data: {
        labels,
        datasets: [{
          label,
          data,
          backgroundColor: this.getColors(labels.length)
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: { legend: { position: "bottom" } }
      }
    };

    this.chart = new Chart(this.chartContext, config);
  }

  private getColors(count: number): string[] {
    const palette = ["#2563eb", "#10b981", "#f59e0b", "#ef4444", "#6366f1", "#14b8a6", "#fbbf24", "#ec4899"];
    return Array.from({ length: count }, (_, i) => palette[i % palette.length]);
  }

  private calculateEntityCounts(): void {
    const entities = ["opportunities", "activities", "orders", "quotes", "products", "accounts", "leads"];
    entities.forEach(e => (this.entityCounts[e] = 0));
    this.renderEntityCounts();
  }

  public updateView(_context?: ComponentFramework.Context<IInputs>): void {
    // Optional: react to dataset or property changes
  }

  public getOutputs(): IOutputs {
    return {};
  }

  public destroy(): void {
    if (this.chart) {
      try {
        this.chart.destroy();
      } catch (e) {
        console.warn("Error destroying chart on destroy:", e);
      }
    }
  }

  private escapeHtml(value: string): string {
    return value
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#039;");
  }
}
