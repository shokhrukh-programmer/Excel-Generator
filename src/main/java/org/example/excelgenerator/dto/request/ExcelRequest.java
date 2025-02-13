package org.example.excelgenerator.dto.request;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

@Data
@AllArgsConstructor
@NoArgsConstructor
public class ExcelRequest {
    private String applicationDate;
    private String entryDate;
    private String clientAccount;
    private String loanPurpose;
    private String creditAmount;
    private String creditDuration;
    private String gracePeriod;
    private String interestRate;
    private String lendingMethod;
    private String fundingSource;
    private String loanDisbursementMethod;
    private String additionalCondition;
    private String legalAddress;
    private String establishmentDate;
    private String foundersAndShares;
    private String loanApprovalDecision;
    private String charterCapitalAmount;
    private String mainActivity;
    private String activeLoanCount;
    private String activeLoanBalance;
    private String overduePrincipalAndInterest;
    private String noLegalProceedingsOrOffBalanceLoans;
    private String katmScoreAbove200;
    private String noUnsatisfactoryLoansInAllBanks;
    private String relatedBusinessEntities;
    private String noOverdueDebtsInBRB;
    private String overdueMoreThan30Days;
    private String overdueMoreThan60Days;
    private String overdueMoreThan90Days;
    private String overdueMoreThan90DaysLast12Months;
    private String contractAmount;
    private String remainingCredit;
    private String purpose;
    private String duration;
    private String overdueScheduledAmount;
    private String overdueInterestAmount;
    private String availableCollateral;
    private String periodOne;
    private String last12MonthsTurnover;
    private String periodTwo;
    private String annualBalance2024;
    private String profitOrLoss;
    private String ownWorkingCapital;
    private String mibUz;
    private String secondRegistryDebt;
    private String clientRevenue;
    private String hasAccountInBRBBank;
    private String borrowerCreditLoad;
    private String debtLoadIndicator;
    private String uninterruptedAccountReceiptsLast12Months;
    private String address;
    private String name;
    private String owner;
    private String ownershipDocument;
    private String cadastralExtract;
    private String noRegisteredResidents;
    private String restrictionInfo;
    private String independentAppraisalPrice;
    private String eValuationPrice;
    private String bankAppraisalPrice;
    private String ownerConsent;
    private String address1;
    private String name1;
    private String owner1;
    private String ownershipDocument1;
    private String cadastralExtract1;
    private String noRegisteredResidents1;
    private String restrictionInfo1;
    private String independentAppraisalPrice1;
    private String eValuationPrice1;
    private String bankAppraisalPrice1;
    private String ownerConsent1;
    private String address2;
    private String name2;
    private String owner2;
    private String ownershipDocument2;
    private String cadastralExtract2;
    private String noRegisteredResidents2;
    private String restrictionInfo2;
    private String independentAppraisalPrice2;
    private String eValuationPrice2;
    private String bankAppraisalPrice2;
    private String ownerConsent2;
    private String insuranceCompanyName;
    private String financialStability;
    private String insuranceAmount;
    private String totalGuaranteeAmount;
    private String propertyGuaranteeAmount;
    private String nonPropertyGuaranteeAmount;
    private String bankConclusion;
    private String lawyerConclusion;
    private String committeeLetter;
    private String underwritingLeadManager;
    private String underwritingDepartmentHead;
    private String financeDepartmentDirector;
}
