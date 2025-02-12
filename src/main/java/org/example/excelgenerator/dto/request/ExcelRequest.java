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
}
