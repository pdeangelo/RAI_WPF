using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Collections;

    /// <summary>
    /// Defines the various application constants values
    /// </summary>
    public static class AppConstants
    {
        //Seperator
        public const string Seperator = ",";

        //Success Css
        public const string SuccessCss = "labelGreen";
        public const string SuccessCssNew = "successMessageGlobal";
        //Error Css
        public const string ErrorCss = "labelError";
        public const string ErrorCssNew = "errorMessage";

        //Bold Css
        public const string BoldCss = " labelBold";

        //New Record 
        public const int NewRecord = -1;

        //All Record 
        public const int AllRecord = -1;

        //All Record including historic offshore
        public const int AllRecordInclHistOff = -2;

        //NOT FOUND
        public const int NotFound = -1;

        //FAILURE
        public const int FAILURE = -1;

        //NULL DATE
        public const string NULLDATE = "01/01/1900";

        //NOT A NUMBER
        public const Double ERROR = 99999999999.99;
        //NOT A NUMBER
        public const Double NA = -99999999999.99;

        public const Decimal IRRCONSTANT = 0.90M;

        //NULL DATE
        public const string Price = "Price";
        //NULL DATE
        public const string CashFlow = "CashFlow";

        //Blank VALUE
        public const float BLANKVALUE = 99999999999.99F;

        public const string Convertible = "Convertible";

        public const string Bonds = "Bonds";


        //Drop Down box Select All Values
        public static class SelectAll
        {
            public const string Text = "--Select--";
            public const int Value = 0;
        }

        //Drop Down box Select All Values
        public static class All
        {
            public const string Text = "--All--";
            public const int Value = -1;
        }
        public static class AllInclHistOn
        {
            public const string Text = "--All (Incl Hist OnShore)--";
            public const int Value = -2;
        }

        //Constants for the TextBox Attributes
        public static class TextBoxAttributes
        {
            public const string DBValue = "DBValue";
            public const string OldValue = "OldValue";

        }

        /// <summary>
        /// Static class for TransactionType
        /// </summary>
        public static class ReconTransactionType
        {
            public const string Buy = "Buy";
            public const string Paydown = "Paydown";
            public const string PaydownRefinance = "Paydown - Refinance";
            public const string PaydownRefinanceWithFundTransfer = "Paydown - Refinance With Fund Transfer";
            public const string Sell = "Sell";
            public const string Interest = "Interest";
            public const string InterestRefinance = "Interest - Refinance";
            public const string InterestRefinanceWithFundTransfer = "Interest - Refinance With Fund Transfer";
            public const string Dividend = "Dividend";
            public const string BuyInternalTransfer = "Buy - Internal Transfer";
            public const string BuyFundTransfer = "Buy - Fund Transfer";
            public const string BuyRestructuring = "Buy - Restructuring";
            public const string BuyDebtToEquityConversion = "Buy - Debt to Equity Conversion";
            public const string BuyRefinance = "Buy - Refinance";
            public const string BuyCLOTransfer = "Buy - CLO Transfer";
            public const string BuyRestructuringWithFundTransfer = "Buy - Restructuring With Fund Transfer";
            public const string BuyRefinanceWithFundTransfer = "Buy - Refinance With Fund Transfer";
            public const string BuyLeverageTransfer = "Buy - Leverage Transfer";
            public const string SellInternalTransfer = "Sell - Internal Transfer";
            public const string SellFundTransfer = "Sell - Fund Transfer";
            public const string SellRestructuring = "Sell - Restructuring";
            public const string SellDebtToEquityConversion = "Sell - Debt to Equity Conversion";
            public const string SellRefinance = "Sell - Refinance";
            public const string SellCLOTransfer = "Sell - CLO Transfer";
            public const string SellRestructuringWithFundTransfer = "Sell - Restructuring With Fund Transfer";
            public const string SellRefinanceWithFundTransfer = "Sell - Refinance With Fund Transfer";
            public const string SellLeverageTransfer = "Sell - Leverage Transfer";
        }

        /// <summary>
        /// Static class for MiscType
        /// </summary>
        public static class MiscType
        {
            public const string Role = "Role";
            public const string Case = "Case";
            public const string Index = "Index";
            public const string OfficeLocation = "Office Location";
            public const string TransactionType = "Transaction Type";
            public const string GenevaTransactionType = "Geneva Transaction Type";
            public const string SecurityType = "Security Type";
            public const string GenevaTypes = "Geneva Types";
            public const string Region = "Region";
            public const string IndustryGroup = "Industry Group";
            public const string Industry = "Industry";
            public const string Sectors = "Sectors";
            public const string InvestmentType = "Investment Type";
            public const string ValuationMethod = "Valuation Method";
            public const string AssetDescription = "Asset Description";
            public const string Custody = "Custody";
            public const string Liquidity = "Liquidity";
            public const string Strategy = "Strategy";
            public const string Tenor = "Tenor";
            public const string PrincipalPercent = "Principal Percent";
            public const string BloombergTenors = "Bloomberg Tenors";
            public const string InterestRateBips = "Interest Rate Bips";
            public const string EquityPerformance = "Equity Performance";
            public const string Timings = "Timings";
            public const string FXRatePercent = "FX Rate Percent";
            public const string AssetType = "Asset Type";
            public const string ReconciliationComments = "Reconciliation Comments";
            public const string FinanceGroup = "Finance Group";
            public const string WSOTransactionType = "WSO Transaction Type";
            public const string CCPIndustry = "CCP Industry";
            public const string CountryatRisk = "Country at Risk";
            public const string MarketingStrategy = "Marketing Strategy";
            public const string Paultest = "Paul test";
            public const string PortfolioClassification = "Portfolio Classification";
            public const string InvestorFlowTranType = "InvestorFlowTranType";
            public const string InvestorType = "InvestorType";
            public const string InvestorTypeDetail = "InvestorTypeDetail";
            public const string Advisor = "Advisor";
            public const string LPGP = "LPGP";
            public const string LPGPDetail = "LPGPDetail";
            public const string PaulTestIndustry = "Paul Test Industry";
            public const string FIBreakup = "FI Breakup";
            public const string Realized_UnrealizedGrp = "Realized_UnrealizedGrp";
            public const string YieldReportGroup1 = "YieldReportGroup1";
            public const string YieldReportGroup2 = "YieldReportGroup2";
            public const string PipelineStatus1 = "Pipeline Status 1";
            public const string PipelineStatus2 = "Pipeline Status 2";
            public const string PipelinePriority = "Pipeline Priority";
            public const string PipelineSourcing = "Pipeline Sourcing";
            public const string PipelineTransactionType = "Pipeline Transaction Type";
            public const string PipelineSecurityType = "Pipeline Security Type";
            public const string PipelinePublic = "Pipeline Public";
            public const string PipelinePrimary = "Pipeline Primary";
            public const string PipelineConfiSigned = "Pipeline Confi Signed";
            public const string PipelineWentToIC = "Pipeline Went To IC";
            public const string PipelineLiquid = "Pipeline Liquid";

            public const string PipelineReasonForDecline = "Pipeline Reason For Decline";
            public const string PipelineCountry = "Pipeline Country";
            public const string PipelineRegion = "Pipeline Region";
            public const string PipelineStrategy = "Pipeline Strategy";
            public const string PipelineIndustry = "Pipeline Industry";
            public const string PipelineSector = "Pipeline Sector";
            public const string PipelineProvince = "Pipeline Province";
            public const string PipelineCity = "Pipeline City";
            public const string PipelineSubIndustry = "Pipeline SubIndustry";

            public const string PipelineState = "Pipeline State";
            public const string PipelineAssetType = "Pipeline Asset Type";
            public const string PipelineExitDifficulty = "Pipeline Exit Difficulty";
            public const string PipelineLeveragability = "Pipeline Leveragability";
            public const string PipelinePrimaryUseOfRealEstate = "Pipeline Primary Use Of Real Estate";

        }

        /// <summary>
        /// Hashtable to store Pipeline status1 values
        /// </summary>
        /// <returns></returns>
        public static Hashtable GetPipelineStatus1Hashtable()
        {

            Hashtable PipelineStatus1 = new Hashtable();
            PipelineStatus1.Add("--Select--", 0);
            PipelineStatus1.Add("Approved", 11793);
            PipelineStatus1.Add("Declined", 11794);
            PipelineStatus1.Add("Monitor", 11795);
            PipelineStatus1.Add("Active", 11967);
            PipelineStatus1.Add("", 0);

            return PipelineStatus1;
        }

        /// <summary>
        /// Hashtable to store Pipeline status2 values
        /// </summary>
        /// <returns></returns>
        public static Hashtable GetPipelineStatus2Hashtable()
        {

            Hashtable PipelineStatus2 = new Hashtable();
            PipelineStatus2.Add("--Select--", 0);
            PipelineStatus2.Add("PE and LTV", 11932);
            PipelineStatus2.Add("PE", 11933);
            PipelineStatus2.Add("LTV", 11934);
            PipelineStatus2.Add("LTV and PE 2014", 12000);
            PipelineStatus2.Add("PE 2014", 12001);
            PipelineStatus2.Add("LTV 2014", 12002);
            

            //PipelineStatus2.Add("Accumulate", 11796);
            //PipelineStatus2.Add("Monitor", 11797);
            //PipelineStatus2.Add("Purchased", 11798);
            //PipelineStatus2.Add("IC Review", 11799);
            
            //PipelineStatus2.Add("PE", 11933);
            //PipelineStatus2.Add("LTV", 11934);
            PipelineStatus2.Add("", 0);

            return PipelineStatus2;
        }


        /// <summary>
        /// Hashtable to store Pipeline Strategy values
        /// </summary>
        /// <returns></returns>
        public static Hashtable GetPipelineStrategyHashtable()
        {

            Hashtable PipelineStrategy = new Hashtable();
            PipelineStrategy.Add("--Select--", 0);
            PipelineStrategy.Add("Equity", 11859);
            PipelineStrategy.Add("Pools/Portfolios", 11860);
            PipelineStrategy.Add("Primary/Direct Lending/ Rescue Debt", 11861);
            PipelineStrategy.Add("Restructuring", 11862);
            PipelineStrategy.Add("Secondary Debt / High Yield", 11863);
            PipelineStrategy.Add("", 0);

            return PipelineStrategy;
        }


        /// <summary>
        /// Hashtable to store PipelinePriority values
        /// </summary>
        /// <returns></returns>
        public static Hashtable GetPipelinePriorityHashtable()
        {

            Hashtable PipelinePriority = new Hashtable();
            PipelinePriority.Add("--Select--", 0);
            PipelinePriority.Add("1 - High", 11800);
            PipelinePriority.Add("2 - Medium High", 11801);
            PipelinePriority.Add("4 - Low", 11802);
            PipelinePriority.Add("5 - Watchlist", 11911);
            PipelinePriority.Add("3 - Medium Low", 11912);
            PipelinePriority.Add("", 0);

            return PipelinePriority;
        }

        /// <summary>
        /// Hashtable to store Pipeline Liquid values
        /// </summary>
        /// <returns></returns>
        public static Hashtable GetPipelineLiquidHashtable()
        {

            Hashtable PipelineLiquid = new Hashtable();
            PipelineLiquid.Add("--Select--", 0);
            PipelineLiquid.Add("Y", 11827);
            PipelineLiquid.Add("N", 11828);
            PipelineLiquid.Add("", 0);

            return PipelineLiquid;
        }

        /// <summary>
        /// Hashtable to store PipelinePublicSec values
        /// </summary>
        /// <returns></returns>
        public static Hashtable GetPipelinePublicSecHashtable()
        {

            Hashtable PipelinePublicSec = new Hashtable();
            PipelinePublicSec.Add("--Select--", 0);
            PipelinePublicSec.Add("Y", 11819);
            PipelinePublicSec.Add("N", 11820);
            PipelinePublicSec.Add("", 0);

            return PipelinePublicSec;
        }
    }
public class ReturnResult
{
    #region Private Members
    /// <summary>
    /// Define private ControlValidation members
    /// </summary>
    private Boolean _status = false;
    private string _message = "";
    #endregion

    #region Constructor
    /// <summary>
    /// Default Constructor.  
    /// </summary>
    public ReturnResult()
    {
        _status = false;
        _message = "";
    }
    #endregion

    #region Public Property
    /// <summary>
    /// Gets the Status.
    /// </summary>
    public Boolean Status
    {
        get { return _status; }
        set { _status = value; }
    }
    /// <summary>
    /// Gets/sets the Error Message.
    /// </summary>
    public string Message
    {
        get { return _message; }
        set { _message = value; }
    }
    #endregion

}