import { IInputs, IOutputs } from "./generated/ManifestTypes";
// import { config } from "./config";

interface Lead {
  id: string;
  name: string;
  title: string;
  date: string;
  status_reason: string;
  label_reason: string;

}

export class caseKanban20 implements ComponentFramework.StandardControl<IInputs, IOutputs> {
  private context: ComponentFramework.Context<IInputs>;
  private notifyOutputChanged: () => void;
  private container: HTMLDivElement;
  private searchQuery = "";
  private statusReasonLabels: Record<number, string> = {
    295290001: "Notification of Issue",
    295290002: "At Risk - Monitoring",
    295290003: "Apprentice Check In",
    295290007: "Employer Check In",
    295290004: "Changing Employer",
    295290005: "Looking for New Employer",
    295290006: "Employer Looking for Apprentice",
    295290008: "AASN Contact Needed"
  };
  private DevActivelabel_reason: Record<number, string> = {
    1: "In Progress",
    100000000: "On Hold",
    522770001: "Waiting for STS Response",
    522770002: "Monitor & Update STS",
    522770003: "Suspension Injury",
    522770004: "Suspension Personal",
    522770005: "Extension Required"
  };  private ProdActivelabel_reason: Record<number, string> = {
    1: "In Progress",
    2: "On Hold",
    3: "Waiting for STS Response",
    4: "Monitor & Update STS",
    295290001: "Suspension Injury",
    295290002: "Suspension Personal",
    295290003: "Extension Required"
  };

  private statusReasonValues: Record<string, number> = {
    "Notification of Issue": 295290001,
    "At Risk - Monitoring": 295290002,
    "Apprentice Check In": 295290003,
    "Employer Check In": 295290007,
    "Changing Employer": 295290004,
    "Looking for New Employer": 295290005,
    "Employer Looking for Apprentice": 295290006,
    "AASN Contact Needed": 295290008
  };


  private fetchColumns = (): Promise<string[]> => {
    return new Promise((resolve) => {
      const customOrder = [295290001, 295290002, 295290003, 295290007, 295290004, 295290005, 295290006, 295290008];
      setTimeout(() => {
        resolve(customOrder.map(key => this.statusReasonLabels[key]));
      }, 1000);
    });
  };


  private fetchData = async (): Promise<Lead[]> => {

    const leads: Lead[] = []

    try {
      let entityName = ""; // Define entity name dynamically
      // Determine which FetchXML query to use based on this.baseUrl
      let fetchXML = "";
      if (this.baseUrl === "https://org7f2e3f28.crm8.dynamics.com") {
        fetchXML = `
             <fetch>
                 <entity name="sam_incident">
                     <attribute name="sam_incidentid" />
                     <attribute name="sam_title" />
                     <attribute name="sam_followupby" />
                     <attribute name="sam_casetypecode" />
                     <attribute name="statuscode" />
        <filter>
                            <condition attribute="statecode" operator="eq" value="0" />
        </filter>
                     <link-entity name="contact" from="contactid" to="sam_customerid" link-type="outer" alias="contact">
                         <attribute name="contactid" />
                         <attribute name="fullname" />
                        
                     </link-entity>

                     <link-entity name="account" from="accountid" to="sam_customerid" link-type="outer" alias="account">
                         <attribute name="accountid" />
                         <attribute name="name" />
                        
                     </link-entity>
                 </entity>
             </fetch>`;
        entityName = "sam_incident"; // Set entity name for API call
      } else if (this.baseUrl === "https://cti.crm6.dynamics.com") {
        fetchXML = `
             <fetch>
                 <entity name="incident">
                     <attribute name="incidentid" />
                     <attribute name="title" />
                     <attribute name="followupby" />
                     <attribute name="casetypecode" />
                     <attribute name="statuscode" />
                        <filter>
                           <condition attribute="statecode" operator="eq" value="0" />
                      </filter>
                     <link-entity name="contact" from="contactid" to="customerid" link-type="outer" alias="contact">
                         <attribute name="contactid" />
                         <attribute name="fullname" />
                        
                     </link-entity>
                    <link-entity name="account" from="accountid" to="customerid" link-type="outer" alias="account">
                         <attribute name="accountid" />
                         <attribute name="name" />
                        
                     </link-entity>

                 </entity>
             </fetch>`;
        entityName = "incident"; // Set entity name for API call
      } else {
        console.error("Base URL does not match any known organization.");
        return leads;
      }


      // Make the API call using the dynamically set entity name
      const result = await this.context.webAPI.retrieveMultipleRecords(entityName, `?fetchXml=${encodeURIComponent(fetchXML)}`);

      result.entities.forEach((entity: ComponentFramework.WebApi.Entity) => {
        // Ensure a valid Date object is always assigned
        const date = entity.sam_followupby
          ? new Date(entity.sam_followupby)
          : new Date(entity.followupby);

        const formattedDate = `${date.getDate().toString().padStart(2, '0')}-${(date.getMonth() + 1).toString().padStart(2, '0')}-${date.getFullYear()}`;
        leads.push({
          id: entity.incidentid || entity.sam_incidentid,
          name: entity["contact.fullname"] || entity["account.name"] || "",
          date: formattedDate || "",
          status_reason: this.statusReasonLabels[entity.sam_casetypecode] || this.statusReasonLabels[entity.casetypecode] || "",
          title: entity["sam_title"] || entity["title"] || "",
          label_reason: this.DevActivelabel_reason[entity.statuscode] || this.ProdActivelabel_reason[entity.statuscode]
        });

      });
      // Sort leads by date (oldest first)
      leads.sort((a, b) => {
        const dateA = new Date(a.date.split("-").reverse().join("-"));
        const dateB = new Date(b.date.split("-").reverse().join("-"));
        return dateA.getTime() - dateB.getTime(); // Oldest first
      });

      return leads;

    } catch (error) {
      return leads;

    }

  };

  private baseUrl: string = window.location.origin;

  private createKanbanCard = (lead: Lead): HTMLDivElement => {
    const card = document.createElement("div");
    card.className = "kanban-card";
    card.setAttribute("draggable", "true");
    card.setAttribute("data-id", lead.id);

    card.innerHTML = `
    <div class="box">
    <h2 class="status-reason">${lead.label_reason}</h2>
    <div class="content-box">
    <h3 class="userName">${lead.title}</h3>
      <div class="details">
        <span>${lead.name} </span>
        
        <span>Follow Up: ${lead.date}</span></div>
      </div>
    </div>  
    </div>
  `;

    card.addEventListener("dragstart", this.handleDragStart);
    card.addEventListener("click", this.handleCardClick);
    return card;
  };

  private handleDragStart = (event: DragEvent): void => {
    const target = event.target as HTMLDivElement;
    const leadId = target.getAttribute("data-id");
    event.dataTransfer?.setData("text/plain", leadId || "");
  };

  private handleCardClick = (event: MouseEvent): void => {
    const target = event.currentTarget as HTMLDivElement;
    const leadId = target.getAttribute("data-id");

    if (!leadId) {
      console.warn("No leadId found on clicked card.");
      return;
    }

    // Determine entity name based on environment
    const entityName = this.baseUrl.includes("https://cti.crm6.dynamics.com") ? "incident" : "sam_incident";

    const entityFormOptions = {
      entityName: entityName, // Use "incident" for Prod, "sam_incident" for Dev
      entityId: leadId
    };

    console.log("Opening form with entity:", entityFormOptions);

    if (this.context.navigation && this.context.navigation.openForm) {
      this.context.navigation.openForm(entityFormOptions)
        .then(() => {
          console.log("Form opened successfully.");
        })
        .catch((error: any) => {
          console.error("Error opening form:", error);
        });
    } else {
      console.error("Navigation context not available. Cannot open form.");
    }
  };

  private handleDrop = async (event: DragEvent): Promise<void> => {
    event.preventDefault();

    const leadId = event.dataTransfer?.getData("text/plain");
    const targetColumn = (event.target as HTMLElement).closest(".kanban-column");
    const newStatus = targetColumn?.getAttribute("data-status");

    if (leadId && newStatus) {
      const newStatusReason = this.statusReasonValues[newStatus];

      if (newStatusReason === undefined) {
        console.error(`Status reason not found for column label: ${newStatus}`);
        return;
      }

      try {
        // Determine the entity name based on baseUrl
        const entityName = this.baseUrl === "https://cti.crm6.dynamics.com" ? "incident" : "sam_incident";

        await this.context.webAPI.updateRecord(entityName, leadId, { sam_casetypecode: newStatusReason });
        this.updateKanbanBoard();
        const card = document.querySelector(`.kanban-card[data-id="${leadId}"]`) as HTMLDivElement;
        if (card) {
          const oldColumnBody = card.closest(".kanban-column-body") as HTMLDivElement;
          if (oldColumnBody && oldColumnBody.contains(card)) {
            oldColumnBody.removeChild(card);
          }

          const newColumnBody = targetColumn?.querySelector(".kanban-column-body") as HTMLDivElement;
          if (newColumnBody) {
            const rect = newColumnBody.getBoundingClientRect();
            const offsetY = event.clientY - rect.top;
            const cards = Array.from(newColumnBody.querySelectorAll(".kanban-card"));

            let insertIndex = cards.findIndex(card => {
              const cardRect = card.getBoundingClientRect();
              return offsetY < (cardRect.top + cardRect.height / 2 - rect.top);
            });

            if (insertIndex === -1) insertIndex = cards.length;

            if (insertIndex < cards.length) {
              newColumnBody.insertBefore(card, cards[insertIndex]);
            } else {
              newColumnBody.appendChild(card);
            }
          }

          const statusFlag = card.querySelector(".status-flag") as HTMLSpanElement;
          if (statusFlag) {
            statusFlag.textContent = newStatus;
          }
        }
      } catch (error) {
        console.error("Error updating lead status:", error);
      }
    } else {
      console.warn("Lead ID or new status not found.");
    }
  };

  private handleDragOver = (event: DragEvent): void => {
    event.preventDefault();
  };

  private createKanbanColumn = (columnLabel: string): HTMLDivElement => {
    const column = document.createElement("div");
    column.className = "kanban-column";
    column.setAttribute("data-status", columnLabel);

    column.innerHTML = `
      <div class="kanban-column-header">${columnLabel}</div>
      <div class="kanban-column-body"></div>
    `;

    column.addEventListener("drop", this.handleDrop);
    column.addEventListener("dragover", this.handleDragOver);
    return column;
  };

  private createSearchAndSignUpContainer(): HTMLDivElement {
    const container = document.createElement("div");
    container.className = "search-and-sign-up-container";

    // Create and append search box
    const searchBoxContainer = document.createElement("div");
    searchBoxContainer.className = "search-box-container";
    searchBoxContainer.innerHTML = `
        <input type="text" id="search-box" placeholder="Search Here..." value="${this.searchQuery}"/>
    `;
    container.appendChild(searchBoxContainer);

    // Create and append sign up button
    const signUpButton = document.createElement("div");
    signUpButton.className = "sign-up-button-container";
    signUpButton.innerHTML = `
        <button id="sign-up-btn" class="sign-up-btn">New</button>
    `;
    container.appendChild(signUpButton);

    const button = signUpButton.querySelector("#sign-up-btn") as HTMLButtonElement;
    button.onclick = () => {
      this.handleSignUp();
    };

    return container;
  }


  private handleSignUp = (): void => {
    // Determine entity name based on environment
    const entityName = this.baseUrl.includes("https://cti.crm6.dynamics.com") ? "incident" : "sam_incident";

    const entityFormOptions = {
      entityName: entityName, // Use "incident" for Prod, "sam_incident" for Dev
    };

    console.log("Opening sign-up form with entity:", entityFormOptions);

    if (this.context.navigation && this.context.navigation.openForm) {
      this.context.navigation.openForm(entityFormOptions)
        .then(() => {
          console.log("Case  form opened successfully.");
        })
        .catch((error: any) => {
          console.error("Error opening Case form:", error);
        });
    } else {
      console.error("Navigation context not available. Cannot open Case form.");
    }
  };



  private async renderKanbanBoard(): Promise<void> {
    this.container.innerHTML = "";

    const searchBoxContainer = this.createSearchAndSignUpContainer();
    this.container.appendChild(searchBoxContainer);

    const searchBox = searchBoxContainer.querySelector("#search-box") as HTMLInputElement;
    searchBox.addEventListener("input", (event: Event) => {
      this.searchQuery = (event.target as HTMLInputElement).value.toLowerCase();
      this.updateKanbanBoard();
    });

    const columnLabels = await this.fetchColumns();
    const kanbanBoard = document.createElement("div");
    kanbanBoard.className = "kanban-board";

    columnLabels.forEach((label) => {
      kanbanBoard.appendChild(this.createKanbanColumn(label));
    });

    this.container.appendChild(kanbanBoard);

    this.updateKanbanBoard(); // Initial render of the Kanban board
  }

  private async updateKanbanBoard(): Promise<void> {
    const columnLabels = Object.values(this.statusReasonLabels);
    const kanbanBoard = this.container.querySelector(".kanban-board") as HTMLDivElement;

    if (!kanbanBoard) return;

    const leads = await this.fetchData();

    const filteredLeads = leads.filter(lead =>
      lead.name.toLowerCase().includes(this.searchQuery) || lead.title.toLowerCase().includes(this.searchQuery)
    );

    columnLabels.forEach((label) => {
      const column = kanbanBoard.querySelector(`.kanban-column[data-status="${label}"]`);
      const columnBody = column?.querySelector(".kanban-column-body");
      if (columnBody) {
        columnBody.innerHTML = "";
      }
    });

    filteredLeads.forEach((lead) => {
      const card = this.createKanbanCard(lead);
      const column = kanbanBoard.querySelector(`.kanban-column[data-status="${lead.status_reason}"]`);
      const columnBody = column?.querySelector(".kanban-column-body");
      if (columnBody) {
        columnBody.appendChild(card);
      }
    });
  }


  public async init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary, container: HTMLDivElement): Promise<void> {
    this.context = context;
    this.notifyOutputChanged = notifyOutputChanged;
    this.container = container;

    // await this.fetchCountryData();

    this.renderKanbanBoard();
  }

  public updateView(context: ComponentFramework.Context<IInputs>): void {
    // Handle updates to the view
  }

  public getOutputs(): IOutputs {
    return {};
  }

  public destroy(): void {
    // Cleanup if necessary
  }
}

