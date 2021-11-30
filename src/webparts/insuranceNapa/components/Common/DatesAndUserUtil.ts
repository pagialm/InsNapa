import { InsuranceNapaState } from "../InsuranceNapaState";
import { IProposal } from "../IProposal";

interface userItemType {
  stateName: string;
  itemName: string;
}
interface SPUser {
  Title: string;
  Email: string;
  Id: number;
}
class DatesAndUserUtil {
  public static datesArray: userItemType[] = [
    {
      stateName: "targetCompletionDate",
      itemName: "TargetCompletionDate",
    },
    {
      stateName: "bUPRCDate",
      itemName: "BUPRCDate",
    },
    {
      stateName: "businessCaseApprovalDate",
      itemName: "BusinessCaseApprovalDate",
    },
    {
      stateName: "targetBusinessGoLive",
      itemName: "TargetBusinessGoLive",
    },
    {
      stateName: "nAPABriefingDate",
      itemName: "NAPABriefingDate",
    },
    {
      stateName: "targetSubmissionByBusiness",
      itemName: "TargetSubmissionByBusiness",
    },
    {
      stateName: "bIRORegionalHeadReviewDate",
      itemName: "BIRORegionalHeadReviewDate",
    },
    {
      stateName: "cROStatusDate",
      itemName: "CROStatusDate",
    },
    {
      stateName: "pIRDateCompleted",
      itemName: "PIRDateCompleted",
    },
    {
      stateName: "targetDueDate",
      itemName: "TargetDueDate",
    },
  ];
  public static usersArray: userItemType[] = [
    {
      stateName: "applicationCompletedBy",
      itemName: "AppCreatedById",
    },
    { stateName: "sponser", itemName: "SponsorId" },
    { stateName: "tradingBookOwner", itemName: "TradingBookOwnerId" },
    {
      stateName: "workstreamCoordinator",
      itemName: "WorkStreamCoordinatorId",
    },
    {
      stateName: "nAPATeamCoordinators",
      itemName: "NAPATeamCoordinatorsId",
    },
    {
      stateName: "infraAreaApprovedByBUPRC",
      itemName: "InfraAreaApprovedByBUPRCId",
    },
    {
      stateName: "LegalReviewer",
      itemName: "LegalReviewerId",
    },
    {
      stateName: "ITReviewer",
      itemName: "ITReviewerId",
    },
    {
      stateName: "CreditRiskReviwer",
      itemName: "CreditRiskReviwerId",
    },
    {
      stateName: "RegulatoryReportingReviewer",
      itemName: "RegulatoryReportingReviewerId",
    },
    {
      stateName: "TreasuryReviewer",
      itemName: "TreasuryReviewerId",
    },
    {
      stateName: "IRMReviewer",
      itemName: "IRMReviewerId",
    },
    {
      stateName: "FinancialReportingReviewer",
      itemName: "FinancialReportingReviewerId",
    },
    {
      stateName: "FraudRiskReviewer",
      itemName: "FraudRiskReviewerId",
    },
    {
      stateName: "TaxReviewer",
      itemName: "TaxReviewerId",
    },
    {
      stateName: "BusinessCaseApprovalFrom",
      itemName: "BusinessCaseApprovalFromId",
    },
    {
      stateName: "LegalReviewer",
      itemName: "LegalReviewerId",
    },
    {
      stateName: "ComplianceReviwer",
      itemName: "ComplianceReviwerId",
    },
    {
      stateName: "OperationsReviewer",
      itemName: "OperationsReviewerId",
    },
    {
      stateName: "MarketRiskReviewer",
      itemName: "MarketRiskReviewerId",
    },
    {
      stateName: "ProductControlReviewer",
      itemName: "ProductControlReviewerId",
    },
    {
      stateName: "CRMReviewer",
      itemName: "CRMReviewerId",
    },
    {
      stateName: "TreasuryRiskReviewer",
      itemName: "TreasuryRiskReviewerId",
    },
    {
      stateName: "GroupResilienceReviewer",
      itemName: "GroupResilienceReviewerId",
    },
    {
      stateName: "FinancialCrimeReviewer",
      itemName: "FinancialCrimeReviewerId",
    },
    {
      stateName: "ConductRiskReviewer",
      itemName: "ConductRiskReviewerId",
    },
    {
      stateName: "LegalReviewer",
      itemName: "LegalReviewerId",
    },
    {
      stateName: "LegalReviewer",
      itemName: "LegalReviewerId",
    },
    {
      stateName: "BIRORegionalHead",
      itemName: "BIRORegionalHeadId",
    },
    {
      stateName: "CustomerExperienceReviewer",
      itemName: "CustomerExperienceReviewerId",
    },
    {
      stateName: "DistributionReviewer",
      itemName: "DistributionReviewerId",
    },
    {
      stateName: "ReinsuranceReviewer",
      itemName: "ReinsuranceReviewerId",
    },
    {
      stateName: "ProductGovernanceCustodians",
      itemName: "ATTChairId",
    },
  ];
  /**
   * GetDisplayNames
   */
  public static GetDisplayNames(
    proposalItem: IProposal,
    getUserIds: (userIds: number[]) => Promise<SPUser[]>
  ): Promise<any[]> {
    let queryString = "";

    this.usersArray.forEach((user, idx, users) => {
      queryString += `${proposalItem[user.itemName]}`;
      if (idx + 1 < users.length) queryString += ",";
    });

    const dispUserIds = queryString.split(",").map((x) => +x);
    const filteredUserIds = dispUserIds.filter((userId) => {
      if (!isNaN(userId)) return userId;
    });
    // console.log(queryString);
    const userFields = [];
    return getUserIds(filteredUserIds).then((items) => {
      this.usersArray.forEach((userItem) => {
        if (Array.isArray(proposalItem[userItem.itemName])) {
          const _userFields = [];
          proposalItem[userItem.itemName].forEach((userId) => {
            const singleUser = items.filter((itm) => {
              if (itm.Id === userId) return itm.Title;
            });
            if (singleUser.length > 0) _userFields.push(singleUser[0].Title);
          });
          const t = {};
          t[userItem["stateName"]] = _userFields;
          userFields.push(t);
        } else {
          const t = {};
          const singleUser = items.filter((itm) => {
            if (itm.Id === proposalItem[userItem.itemName]) return itm.Title;
          });
          t[userItem["stateName"]] =
            singleUser.length > 0 ? singleUser[0].Title : "";
          userFields.push(t);
          // console.log(t, userFields);
        }
      });
      return userFields as unknown as Promise<any[]>;
    });
  }
  /**
   * GetDates
   */
  public static GetDates(): any[] {
    return this.datesArray;
  }
}

export default DatesAndUserUtil;
