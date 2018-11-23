# Lookup values for financial dimensions to Excel templates
The goal is to automate the addition of lookup for custom financial dimensions within excel templates.

From Visual Studio: **Dynamics 365 > Add-ins > Add financial dimensions lookup for OData**

![new_form](https://github.com/anderson-joyle/D365FO-ExcelFinancialDimensionLookups/blob/master/Addin.png)

## How to
1. Make sure you have added new dimension fields into **DimensionCombinationEntity** and **DimensionSetEntity** data entities using "Add financial dimensions for OData" native add-in.
2. Open add-in via Visual Studio menu bar: **Dynamics 365 > Add-ins > Add financial dimensions lookup for OData**
3. Select your customizations model.
4. Hit **Apply** button.

Addin code will try to find extension for each **DimensionCombinationEntity** and **DimensionSetEntity** data entities within the selected model.
Once they are found, a relationship will be created for each new field, using this [table map](https://docs.microsoft.com/en-us/dynamics365/unified-operations/dev-itpro/financial/add-dimensions-excel-templates) as a guide.

> Name convertion is really important here. Before to execute this addin, take a look on how your financial dimensions has been named. It will try to match field name with dimension type (Agreements, Bank accounts, Departments, etc) by removing all blank spaces and singluar/plural naming i.e. "Bank accounts" will be checked against <i>bankaccount</i> and <i>bankaccounts</i> values.

## Limitations
* <i>Fixed assets (Russia)</i> dimension is not being checked.
