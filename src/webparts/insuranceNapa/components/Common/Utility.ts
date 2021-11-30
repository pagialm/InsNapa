import * as React from "react";
interface ICustomInfraObject {
  comment: string;
  approvalDate: string;
  approvedBy: string;
}

class Utility {
  public static GetInfraObject = (infaItem: string): ICustomInfraObject => {
    console.log(infaItem);
    const infraObj: ICustomInfraObject = {
      comment: "",
      approvedBy: "",
      approvalDate: "",
    };

    let infraCommentField,
      infraApprovalDate,
      infraApprovalBy = "";
    switch (infaItem) {
      case "Legal Risk": {
        infraCommentField = "RemoveLegalApprovalComment";
        infraApprovalDate = "LegalApprovalDate";
        infraApprovalBy = "LegalApprovedBy0";
        break;
      }
      case "Compliance": {
        infraCommentField = "RemoveComplianceApprovalComment";
        infraApprovalDate = "ComplianceApprovalDate";
        infraApprovalBy = "ComplianceApprovedBy0";
        break;
      }
      case "Valuations": {
        infraCommentField = "RemoveConductRiskApprovalComment";
        infraApprovalDate = "ConductRiskApprovalDate";
        infraApprovalBy = "ConductRiskApprovedBy0";
        break;
      }
      case "CRO": {
        infraCommentField = "RemoveITApprovalComment";
        infraApprovalDate = "ITApprovalDate";
        infraApprovalBy = "ITApprovedBy0";
        break;
      }
      case "Operations": {
        infraCommentField = "RemoveOperationsApprovalComment";
        infraApprovalDate = "OperationsApprovalDate";
        infraApprovalBy = "OperationsApprovedBy0";
        break;
      }
      case "Data Privacy": {
        infraCommentField = "RemoveCreditRiskApprovalComment";
        infraApprovalDate = "CreditRiskApprovalDate";
        infraApprovalBy = "CreditRiskApprovedBy0";
        break;
      }
      case "Information Security Risk and Cyber Risk": {
        infraCommentField = "RemoveMarketRiskApprovalComment";
        infraApprovalDate = "MarketRiskApprovalDate";
        infraApprovalBy = "MarketRiskApprovedBy0";
        break;
      }
      case "Tax Risk": {
        infraCommentField = "RemoveTaxApprovalComment";
        infraApprovalDate = "TaxApprovalDate";
        infraApprovalBy = "TaxApprovedBy0";
        break;
      }
      case "Head of Actuarial and Statutory Actuary": {
        infraCommentField = "RemoveProductControlApprovalComm";
        infraApprovalDate = "ProductControlApprovalDate";
        infraApprovalBy = "ProductControlApprovedBy0";
        break;
      }
      case "Marketing and Communications": {
        infraCommentField = "RemoveRegulatoryApprovalComment";
        infraApprovalDate = "RegulatoryReportingApprovalDate";
        infraApprovalBy = "RegulatoryReportingApprovedBy0";
        break;
      }
      case "Financial Reporting/ Control Risk": {
        infraCommentField = "RemoveFinancialReportingApproval0";
        infraApprovalDate = "FinancialReportingApprovalDate";
        infraApprovalBy = "FinancialReportingApprovedBy0";
        break;
      }
      case "Supplier Risk": {
        infraCommentField = "RemoveTreasuryApprovalComment";
        infraApprovalDate = "TreasuryApprovalDate";
        infraApprovalBy = "TreasuryApprovedBy0";
        break;
      }
      case "Financial & Insurance Risk": {
        infraCommentField = "RemoveTreasuryRiskApprovalCommen";
        infraApprovalDate = "TreasuryRiskApprovalDate";
        infraApprovalBy = "TreasuryRiskApprovedBy0";
        break;
      }
      case "Technology Risk": {
        infraCommentField = "RemoveIRMApprovalComment";
        infraApprovalDate = "IRMApprovalDate";
        infraApprovalBy = "IRMApprovedBy0";
        break;
      }
      case "Business Continuity Risk": {
        infraCommentField = "RemoveGroupResilienceApprovalCom";
        infraApprovalDate = "GroupResilienceApprovalDate";
        infraApprovalBy = "GroupResilienceApprovedBy0";
        break;
      }
      case "Fraud Risk": {
        infraCommentField = "RemoveFraudRiskApprovalComment";
        infraApprovalDate = "FraudRiskApprovalDate";
        infraApprovalBy = "FraudRiskApprovedBy0";
        break;
      }
      case "Finance": {
        infraCommentField = "RemoveFinancialCrimeApprovalComm";
        infraApprovalDate = "FinancialCrimeApprovalDate";
        infraApprovalBy = "FinancialCrimeApprovedBy0";
        break;
      }
      case "RBB CVM": {
        infraCommentField = "RemoveORMApprovalComment";
        infraApprovalDate = "CRMApprovalDate";
        infraApprovalBy = "ORMApprovedBy";
        break;
      }
      case "Reinsurance": {
        infraCommentField = "RemoveReinsuranceApprovalComm";
        infraApprovalDate = "ReinsuranceApprovalDate";
        infraApprovalBy = "ReinsuranceApprovedBy";
        break;
      }
      case "Customer Experience": {
        infraCommentField = "RemoveCustomerExperienceApprovalComm";
        infraApprovalDate = "CustomerExperienceeApprovalDate";
        infraApprovalBy = "CustomerExperienceApprovedBy";
        break;
      }
      case "Distribution": {
        infraCommentField = "RemoveDistributionApprovalComm";
        infraApprovalDate = "DistributionApprovalDate";
        infraApprovalBy = "DistributionApprovedBy";
        break;
      }
      case "Financial Crime": {
        infraCommentField = "RemoveFinancialCrimeApprovalComm";
        infraApprovalDate = "FinancialCrimeApprovalDate";
        infraApprovalBy = "FinancialCrimeApprovedBy0";
        break;
      }
      default: {
        break;
      }
    }
    infraObj["comment"] = infraCommentField;
    infraObj["approvalDate"] = infraApprovalDate;
    infraObj["approvedBy"] = infraApprovalBy;
    return infraObj;
  };

  /**
   * GetMenuItems
   */
  public static GetMenuItems(): any[] {
    const _PERMENANT = "permanant";
    const stages = [
      { id: 0, title: "Enquiry", selected: false, type: "menu", enabled: true },
      {
        id: 1,
        title: "Proposal",
        selected: false,
        type: "menu",
        enabled: false,
      },
      {
        id: 2,
        title: "Pipeline",
        selected: false,
        type: "menu",
        enabled: false,
      },
      {
        id: 3,
        title: "NPS Pipeline Review",
        selected: false,
        type: "menu",
        enabled: false,
      },
      {
        id: 4,
        title: "Infrastructure Review",
        selected: false,
        type: "stick",
        enabled: false,
      },
      {
        id: 4,
        title: "Infrastructure Review",
        subtile: "CRO",
        selected: false,
        type: "submenu",
        internalName: "IT",
        approved: false,
        enabled: false,
      },
      {
        id: 4,
        title: "Infrastructure Review",
        subtile: "Legal Risk",
        selected: false,
        type: "submenu",
        internalName: "Legal",
        approved: false,
        enabled: false,
      },
      {
        id: 4,
        title: "Infrastructure Review",
        subtile: "Financial Crime",
        selected: false,
        type: "submenu",
        internalName: "FinCrime",
        approved: false,
        enabled: false,
      },
      {
        id: 4,
        title: "Infrastructure Review",
        subtile: "Data Privacy",
        selected: false,
        type: "submenu",
        internalName: "CreditRisk",
        approved: false,
        enabled: false,
      },
      {
        id: 4,
        title: "Infrastructure Review",
        subtile: "Fraud Risk",
        selected: false,
        type: "submenu",
        internalName: "FraudRisk",
        approved: false,
        enabled: false,
      },
      {
        id: 4,
        title: "Infrastructure Review",
        subtile: "Tax Risk",
        selected: false,
        type: "submenu",
        internalName: "Tax",
        approved: false,
        enabled: false,
      },
      {
        id: 4,
        title: "Infrastructure Review",
        subtile: "Information Security Risk and Cyber Risk",
        selected: false,
        type: "submenu",
        internalName: "MarketRisk",
        approved: false,
        enabled: false,
      },
      {
        id: 4,
        title: "Infrastructure Review",
        subtile: "Finance",
        selected: false,
        type: "submenu",
        internalName: "Finance",
        approved: false,
        enabled: false,
      },
      {
        id: 4,
        title: "Infrastructure Review",
        subtile: "Head of Actuarial and Statutory Actuary",
        selected: false,
        type: "submenu",
        internalName: "ProductControl",
        approved: false,
        enabled: false,
      },
      {
        id: 4,
        title: "Infrastructure Review",
        subtile: "Marketing and Communications",
        selected: false,
        type: "submenu",
        internalName: "RegulatoryReporting",
        approved: false,
        enabled: false,
      },
      {
        id: 4,
        title: "Infrastructure Review",
        subtile: "Financial & Insurance Risk",
        selected: false,
        type: "submenu",
        internalName: "Treasury",
        approved: false,
        enabled: false,
      },
      {
        id: 4,
        title: "Infrastructure Review",
        subtile: "Compliance",
        selected: false,
        type: "submenu",
        internalName: "Compliance",
        approved: false,
        enabled: false,
      },
      {
        id: 4,
        title: "Infrastructure Review",
        subtile: "Operations",
        selected: false,
        type: "submenu",
        internalName: "Operations",
        approved: false,
        enabled: false,
      },
      {
        id: 4,
        title: "Infrastructure Review",
        subtile: "Supplier Risk",
        selected: false,
        type: "submenu",
        internalName: "TreasuryRisk",
        approved: false,
        enabled: false,
      },
      {
        id: 4,
        title: "Infrastructure Review",
        subtile: "Financial Reporting/ Control Risk",
        selected: false,
        type: "submenu",
        internalName: "FinTag",
        approved: false,
        enabled: false,
      },
      // {
      //   id: 4,
      //   title: "Infrastructure Review",
      //   subtile: "Risk",
      //   selected: false,
      //   type: "submenu",
      //   approved: false,
      //   enabled: false,
      // },
      {
        id: 4,
        title: "Infrastructure Review",
        subtile: "Technology Risk",
        selected: false,
        type: "submenu",
        internalName: "IRM",
        approved: false,
        enabled: false,
      },
      {
        id: 4,
        title: "Infrastructure Review",
        subtile: "Business Continuity Risk",
        selected: false,
        type: "submenu",
        internalName: "GroupReslience",
        approved: false,
        enabled: false,
      },
      {
        id: 4,
        title: "Infrastructure Review",
        subtile: "RBB CVM",
        selected: false,
        type: "submenu",
        internalName: "CRM",
        approved: false,
        enabled: false,
      },
      {
        id: 4,
        title: "Infrastructure Review",
        subtile: "Valuations",
        selected: false,
        type: "submenu",
        internalName: "ConductRisk",
        approved: false,
        enabled: false,
      },
      {
        id: 4,
        title: "Infrastructure Review",
        subtile: "Reinsurance",
        selected: false,
        type: "submenu",
        internalName: "Reinsurance",
        approved: false,
        enabled: false,
      },
      {
        id: 4,
        title: "Infrastructure Review",
        subtile: "Customer Experience",
        selected: false,
        type: "submenu",
        internalName: "CustomerExperience",
        approved: false,
        enabled: false,
      },
      {
        id: 4,
        title: "Infrastructure Review",
        subtile: "Distribution",
        selected: false,
        type: "submenu",
        internalName: "Distribution",
        approved: false,
        enabled: false,
      },
      {
        id: 5,
        title: "Final NPS Review",
        selected: false,
        type: "menu",
        enabled: false,
      },
      {
        id: 6,
        title: "Chair Review",
        selected: false,
        type: "menu",
        enabled: false,
        hidden: true,
      },
      {
        id: 61,
        title: "Approval to Trade",
        selected: false,
        type: "menu",
        enabled: false,
      },
      {
        id: 62,
        title: "Approved to Trade",
        selected: false,
        type: "menu",
        enabled: false,
        hidden: true,
      },
      {
        id: 7,
        title: "Approval Summary",
        selected: false,
        type: _PERMENANT,
        enabled: false,
      },
      {
        id: 8,
        title: "Other Status",
        selected: false,
        type: _PERMENANT,
        enabled: false,
      },
    ];
    return stages;
  }

  public static GetMenuItemTitle(internalName: string): string {
    let itemTitle = "";
    const menuItems = this.GetMenuItems();
    const filteredMenus = menuItems.filter(
      (fmi) => fmi.internalName === internalName
    );
    if (filteredMenus.length > 0) {
      itemTitle = filteredMenus[0].subtile;
    }

    return itemTitle;
  }
}

export default Utility;
