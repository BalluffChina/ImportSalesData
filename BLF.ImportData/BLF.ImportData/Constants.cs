using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BLF.ImportData
{
    class Constants
    {
        public const string DBName = "BLF_DataCollection";
        public const string AccountTable = "blf_Account";
        public const string AccountMappingTable = "blf_AccountMapping";
        public const string OrderTable = "blf_Order";
        public const string BudgetTable = "blf_Budget";
        public const string SalesTable = "blf_Sales";
        public const string MtlsTable = "blf_Material";
        public const string MtlsPriceTable = "blf_MaterialPrice";
        public const string OpenOrderTable = "blf_OpenOrder";
        public const string ForeignOrderTable = "blf_ForeignOrder";
        public const string ExchangeRatesTable = "blf_ExchangeRates";
        public const int RowLimit = 20000;

        //public const string DBName = "BLF_DataCollection";
        //public const string AccountTable = "blf_Account";
        //public const string AccountMappingTable = "1";
        //public const string OrderTable = "blf_Order";
        //public const string BudgetTable = "3";
        //public const string SalesTable = "4";
        //public const string MtlsTable = "5";
        //public const string MtlsPriceTable = "6";

        public static readonly string GetTables = @"use {0} select name  from sys.tables";
        public static readonly string strInsertAccount = @"INSERT INTO [BLF_DataCollection].[dbo].[blf_Account]
                                                               ([SAP No]
                                                               ,[Customer Name]
                                                               ,[Customer English Name]
                                                               ,[Payment terms]
                                                               ,[O/E/D]
                                                               ,[City]
                                                               ,[Created]
                                                               ,[Region]
                                                               ,[Cust.Class]
                                                               ,[Area Sales Office]
                                                               ,[Sales Engineer]
                                                               ,[Sales Group]
                                                               ,[Province]
                                                               ,[Industry Code]
                                                               ,[Industry Code Description]
                                                               ,[Industry Code1]
                                                               ,[Industry Code1 Description]
                                                               ,[Industry Code3]
                                                               ,[Industry Code4]
                                                               ,[Industry Code5]
                                                               ,[Delete Flag]
                                                               ,[Customer Type]
                                                               ,[Postal Code]
                                                               ,[Customer English Name2]
                                                               ,[City In English]
                                                               ,[Companies To Verify]                         
                                                               ,[First Order Date]
                                                               ,[Operating Organization]
                                                               ,[Vertical Devision]
                                                               ,[Status])
           
                                                         VALUES
                                                               (@SAPNo
                                                               ,@CustomerName
                                                               ,@CustomerEnglishName
                                                               ,@Paymentterms
                                                               ,@OED
                                                               ,@City
                                                               ,convert(datetime,@Created)
                                                               ,@Region
                                                               ,@CustClass
                                                               ,@Office
                                                               ,@SalesEngineer
                                                               ,@SalesGroup
                                                               ,@Province
                                                               ,@IndustryCode
                                                               ,@IndustryCodeDescription
                                                               ,@IndustryCode1
                                                               ,@IndustryCode1Description
                                                               ,@IndustryCode3
                                                               ,@IndustryCode4
                                                               ,@IndustryCode5
                                                               ,@DeleteFlag
                                                               ,@CustomerType
                                                               ,@PostalCode
                                                               ,@CustomerEnglishName2
                                                               ,@CityEn
                                                               ,@CompaniesToVerify
                                                               ,@FirstOrderDate
                                                               ,@OperatingOrganization
                                                               ,@VerticalDevision
                                                               ,@Status
                                                               )";//,[DateAdded],[LastUpdated]

        public static readonly string strInsertBudget = @"INSERT INTO [BLF_DataCollection].[dbo].[blf_Budget]
                                                                   ([Budget]
                                                                   ,[Year]
                                                                   ,[Customer Name]
                                                                   ,[SAP No.]
                                                                   ,[Customer English Name])
                                                             VALUES
                                                                   (@Budget
                                                                   ,@Year
                                                                   ,@CustomerName
                                                                   ,@SAPNo
                                                                   ,@CustomerEnglishName)";

        public static readonly string strInsertAccountMapping = @"INSERT INTO [BLF_DataCollection].[dbo].[blf_AccountMapping]
                                                                   ([SAP NO.]
                                                                   ,[Customer Name]
                                                                   ,[Customer English Name]
                                                                   ,[Industry Code]
                                                                   ,[Industry Code1]
                                                                   ,[Customer Type]
                                                                   ,[Sales Engineer]
                                                                   ,[Area sales Office]
                                                                   ,[Customer Classification])
                                                             VALUES
                                                                   (@SAPNO
                                                                   ,@CustomerName
                                                                   ,@CustomerEnglishName
                                                                   ,@IndustryCode
                                                                   ,@IndustryCode1
                                                                   ,@CustomerType
                                                                   ,@SalesEngineer
                                                                   ,@AreaSalesOffice
                                                                   ,@CustomerClassification)";

        public static readonly string strInsertOrder = @"INSERT INTO [BLF_DataCollection].[dbo].[blf_Order]
                                                                   ([Customer Name]
                                                                   ,[Customer English Name]
                                                                   ,[Sale Engineer]
                                                                   ,[Sale Office]
                                                                   ,[IO Created Year]
                                                                   ,[Calendar Week]
                                                                   ,[Year]
                                                                   ,[Calendar Day]
                                                                   ,[Material ID]
                                                                   ,[Material]
                                                                   ,[Order Code]
                                                                   ,[PlanGrp.1(MD) ID]
                                                                   ,[PlanGrp.1(MD)]
                                                                   ,[PlanGrp.2(MD) ID]
                                                                   ,[PlanGrp.2(MD)]
                                                                   ,[Order Item No.]
                                                                   ,[IO netprc SD(SC)]
                                                                   ,[IO qty]
                                                                   ,[Order No.]
                                                                   ,[SAP No.]
                                                                   ,[Customer Classification]
                                                                   ,[O/E/D]
                                                                   ,[Industry Code]
                                                                   ,[Industry Code1]
                                                                   ,[SL netprc SD (SC)]
                                                                   ,[CustomerType]
                                                                   ,[Region]
                                                                   ,[Payment terms]
                                                                   ,[CreatedAt]
                                                                   ,[City]
                                                                   ,[SL qty]
                                                                   ,[Figure Booking]
                                                                   ,[CRMProjectID]
                                                                   ,[ProjectName]
                                                                   ,[ApplicantSalesOffice]
                                                                   ,[ApplicantSalesEngineer]
                                                                   ,[EndUserName]
                                                                   ,[SalesEngineerByEndUer]
                                                                   ,[SalesOfficeByEndUser]
                                                                   ,[SalesRegionByEndUser]
                                                                   ,[Product Area]
                                                                   ,[Product Group]
                                                                   ,[Order Type]
                                                                   ,[Condition Type])
                                                             VALUES
                                                                   (@CustomerName
                                                                   ,@CustomerEnglishName
                                                                   ,@SaleEngineer
                                                                   ,@SaleOffice
                                                                   ,CONVERT(int, @IOCreatedYear)
                                                                   ,CONVERT(int, @CalendarWeek)
                                                                   ,CONVERT(int, @Year)
                                                                   ,CONVERT(int, @CalendarDay)
                                                                   ,CONVERT(int, @MaterialID)
                                                                   ,@Material
                                                                   ,@OrderCode
                                                                   ,@PlanGrp1MDID
                                                                   ,@PlanGrp1MD
                                                                   ,@PlanGrp2MDID
                                                                   ,@PlanGrp2MD
                                                                   ,@OrderItemNo
                                                                   ,@IOnetprcSD_SC
                                                                   ,CONVERT(int, @IOqty)
                                                                   ,@OrderNo
                                                                   ,CONVERT(int, @SAPNo)
                                                                   ,@CustomerClassification
                                                                   ,@OED
                                                                   ,@IndustryCode
                                                                   ,@IndustryCode1
                                                                   ,@SLnetprcSD_SC
                                                                   ,@CustomerType
                                                                   ,@Region
                                                                   ,@PaymentTerms
                                                                   ,CONVERT(int, @CreatedAt)
                                                                   ,@City
                                                                   ,@SLqty
                                                                   ,@FigureBooking
                                                                   ,@CRMProjectID
                                                                   ,@ProjectName
                                                                   ,@ApplicantSalesOffice
                                                                   ,@ApplicantSalesEngineer
                                                                   ,@EndUserName
                                                                   ,@SalesEngineerByEndUer
                                                                   ,@SalesOfficeByEndUser
                                                                   ,@SalesRegionByEndUser
                                                                   ,@ProductArea
                                                                   ,@ProductGroup
                                                                   ,@OrderType
                                                                   ,@ConditionType
                                                                   )";

        public static readonly string strInsertSales = @"INSERT INTO [BLF_DataCollection].[dbo].[blf_Sales]
                                                                   ([Sale Engineer]
                                                                   ,[Sale Office]
                                                                   ,[Status]
                                                                   ,[OnBoardTime]
                                                                   ,[ExitTime])
                                                             VALUES
                                                                   (@SaleEngineer
                                                                   ,@SaleOffice
                                                                   ,@Status
                                                                   ,@OnBoardTime
                                                                   ,@ExitTime)";

        public static readonly string strInsertMtl = @"INSERT INTO [BLF_DataCollection].[dbo].[blf_Material]
                                                                   ([Material Description]
                                                                   ,[Campaign_C28]
                                                                   ,[Campaign_C26]
                                                                   ,[Campaign_C15]
                                                                   ,[Campaign_C4])
                                                             VALUES
                                                                    (@Mtl
                                                                    ,@C28
                                                                    ,@C26
                                                                    ,@C15
                                                                    ,@C4)";
        public static readonly string strUpdateMtl = @"UPDATE [BLF_DataCollection].[dbo].[blf_Material]
                                                               SET [Campaign_C28] = ''
                                                                  
                                                             WHERE [Material Description] ='as'";

        public static readonly string strInsertMtlPrice = @"INSERT INTO [BLF_DataCollection].[dbo].[blf_MaterialPrice]
                                                                   ([Material]
                                                                   ,[CLP in CNY]
                                                                   ,[GLP in Euro]
                                                                   ,[Product Area]
                                                                   ,[Product Group]
                                                                   ,[New Product]
                                                                   ,[Order Code])
                                                             VALUES
                                                                    (@Material
                                                                    ,@CLPinCNY
                                                                    ,@GLPinEuro
                                                                    ,@ProductArea
                                                                    ,@ProductGroup
                                                                    ,@NewProduct
                                                                    ,@OrderCode
                                                                    )";

        public static readonly string strInsertOpenOrder = @"INSERT INTO [BLF_DataCollection].[dbo].[blf_OpenOrder]
                                                                   ([Calendar Day]
                                                                   ,[Order Number]
                                                                   ,[Order Item Number]
                                                                   ,[Open Value]
                                                                   ,[Period])
                                                             VALUES
                                                                    (@CalendarDay
                                                                    ,@OrderNumber
                                                                    ,@OrderItemNumber
                                                                    ,@OpenValue
                                                                    ,@Period
                                                                    )";

        public static readonly string strInsertOrderUpdate = @"UPDATE [BLF_DataCollection].[dbo].[blf_Order]
                                                               SET [Order Type] = @OrderType
                                                                  
                                                             WHERE [Order No.] = @OrderNo";

        public static readonly string strInsertForeignOrder = @"INSERT INTO [BLF_DataCollection].[dbo].[blf_ForeignOrder]
                                                                   ([Country/Region]
	                                                                ,[Calendar Day]
	                                                                ,[Account No. for Customer]
	                                                                ,[Customer Name Local Language]
	                                                                ,[Country/Region Sold to]
	                                                                ,[Sales Office]
	                                                                ,[Name of Sales Employee]
	                                                                ,[IC0 Industry]
	                                                                ,[IC1 Vertical Industry]
	                                                                ,[IC5 Focus Industry]
	                                                                ,[Postal Code]
	                                                                ,[Customer Name in English]
	                                                                ,[Customer Name in English 2]
	                                                                ,[City in English]
	                                                                ,[Order Number]
	                                                                ,[Invoice no.]
	                                                                ,[Document Type]
	                                                                ,[Material Description]
	                                                                ,[Order Code]
	                                                                ,[Product Group (MD)]
	                                                                ,[Product Area (MD)]
	                                                                ,[IO Gross (SC)]
	                                                                ,[IO qty (Base Unit)]
	                                                                ,[SL Gross (SC)]
	                                                                ,[SL qty (Base Unit)])
                                                             VALUES
                                                                    (@CountryRegion
	                                                                ,CONVERT(int, @CalendarDay)
	                                                                ,@AccountNoForCustomer
	                                                                ,@CustomerNameLocalLanguage
	                                                                ,@CountryRegionSoldTo
	                                                                ,@SalesOffice
	                                                                ,@NameOfSalesEmployee
	                                                                ,@IC0
	                                                                ,@IC1
	                                                                ,@IC5
	                                                                ,@PostalCode
	                                                                ,@CustomerNameInEnglish
	                                                                ,@CustomerNameInEnglish2
	                                                                ,@CityInEnglish
	                                                                ,@OrderNumber
	                                                                ,@InvoiceNo
	                                                                ,@DocumentType
	                                                                ,@MaterialDescription
	                                                                ,@OrderCode
	                                                                ,@ProductGroup
	                                                                ,@ProductArea
	                                                                ,@IOGross_SC
	                                                                ,CONVERT(int, @IOQty_BaseUnit)
	                                                                ,@SLGross_SC
	                                                                ,CONVERT(int, @SLQty_BaseUnit)
                                                                    )";
        public static readonly string strInsertExchangeRates = @"INSERT INTO [BLF_DataCollection].[dbo].[blf_ExchangeRates]
                                                                   ([Country/Region]
                                                                   ,[Rate])
                                                             VALUES
                                                                    (@CountryRegion
                                                                    ,@Rate
                                                                    )";
         
    }
}
