namespace Addin
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using System.ComponentModel.Composition;
    using System.Drawing;
    using Microsoft.Dynamics.Framework.Tools.Extensibility;
    using Microsoft.Dynamics.Framework.Tools.MetaModel.Core;

    using Metadata = Microsoft.Dynamics.AX.Metadata;


    /// <summary>
    /// TODO: Say a few words about what your AddIn is going to do
    /// </summary>
    [Export(typeof(IMainMenu))]
    public class MainMenuAddIn : MainMenuBase
    {
        #region Member variables
        private const string addinName = "Addin";

        private Metadata.Providers.IMetadataProvider metadataProvider = null;
        private Metadata.Service.IMetaModelService metaModelService = null;
        #endregion

        #region Properties
        /// <summary>
        /// Caption for the menu item. This is what users would see in the menu.
        /// </summary>
        public override string Caption
        {
            get
            {
                return AddinResources.MainMenuAddInCaption;
            }
        }

        public Metadata.Providers.IMetadataProvider MetadataProvider
        {
            get
            {
                if (this.metadataProvider == null)
                {
                    this.metadataProvider = DesignMetaModelService.Instance.CurrentMetadataProvider;
                }
                return this.metadataProvider;
            }
        }

        public Metadata.Service.IMetaModelService MetaModelService
        {
            get
            {
                if (this.metaModelService == null)
                {
                    this.metaModelService = DesignMetaModelService.Instance.CurrentMetaModelService;
                }
                return this.metaModelService;
            }
        }

        /// <summary>
        /// Unique name of the add-in
        /// </summary>
        public override string Name
        {
            get
            {
                return MainMenuAddIn.addinName;
            }
        }

        #endregion

        #region Callbacks
        /// <summary>
        /// Called when user clicks on the add-in menu
        /// </summary>
        /// <param name="e">The context of the VS tools and metadata</param>
        public override void OnClick(AddinEventArgs e)
        {
            FormMain formMain = new FormMain(this.MetaModelService.GetModelNames().ToList<string>());
            string modelName = string.Empty;

            if (formMain.ShowDialog() != System.Windows.Forms.DialogResult.OK)
            {
                return;
            }

            modelName = formMain.getModelName();

            if (modelName == string.Empty)
            {
                throw new Exception("Model name cannot be empty.");
            }

            try
            {
                Metadata.MetaModel.AxDataEntityViewExtension dataEntityExtension;

                List<string> entities = this.MetadataProvider.DataEntityViewExtensions.ListObjectsForModel(modelName).ToList<string>().Where<string>(x => x.Contains("DimensionCombinationEntity") || x.Contains("DimensionSetEntity")).ToList<string>();

                string updatedDataEntities = string.Empty;
                string names = string.Empty;

                foreach (var extensionName in entities)
                {
                    bool updated = false;

                    dataEntityExtension = this.MetadataProvider.DataEntityViewExtensions.Read(extensionName);

                    foreach (var field in dataEntityExtension.Fields)
                    {
                        string relatedDataEntity;
                        string nameL = field.Name.ToLower().Replace(" ", "");

                        // List based on https://docs.microsoft.com/en-us/dynamics365/unified-operations/dev-itpro/financial/add-dimensions-excel-templates
                        #region Related entity names
                        switch (nameL)
                        {

                            case "agreement":
                            case "agreements":
                                relatedDataEntity = "DimAttributeAgreementHeaderExt_RUEntity";
                                break;

                            case "bankaccount":
                            case "bankaccounts":
                                relatedDataEntity = "DimAttributeBankAccountTableEntity";
                                break;

                            case "businessunit":
                            case "businessunits":
                                relatedDataEntity = "DimAttributeOMBusinessUnitEntity";
                                break;

                            case "campaign":
                            case "campaigns":
                                relatedDataEntity = "DimAttributeSmmCampaignTableEntity";
                                break;

                            case "cashaccount":
                            case "cashaccounts":
                                relatedDataEntity = "DimAttributeRCashTable_RUEntity";
                                break;

                            case "costcenter":
                            case "costcenters":
                                relatedDataEntity = "DimAttributeOMCostCenterEntity";
                                break;

                            case "customergroup":
                            case "customergroups":
                                relatedDataEntity = "DimAttributeCustGroupEntity";
                                break;

                            case "customer":
                            case "customers":
                                relatedDataEntity = "DimAttributeCustTableEntity";
                                break;

                            case "deferral":
                            case "deferrals":
                                relatedDataEntity = "DimAttributeRDeferralsTable_RUEntity";
                                break;

                            case "department":
                            case "departments":
                                relatedDataEntity = "DimAttributeOMDepartmentEntity";
                                break;

                            case "expenseandincomecode":
                            case "expenseandincomecodes":
                                relatedDataEntity = "DimAttributeRTax25ProfitTable_RUEntity";
                                break;

                            case "expensepurpose":
                            case "expensepurposes":
                                relatedDataEntity = "DimAttributeTrvTravelTxtEntity";
                                break;

                            case "fiscalestablishment":
                            case "fiscalestablishments":
                                relatedDataEntity = "DimAttributeFiscalEstablishment_BREntity";
                                break;

                            case "fixedassetgroup":
                            case "fixedassetgroups":
                                relatedDataEntity = "DimAttributeAssetGroupEntity";
                                break;

                            case "fixedasset":
                            case "fixedassets":
                                relatedDataEntity = "DimAttributeAssetTableEntity";
                                break;

                            case "fund":
                            case "funds":
                                relatedDataEntity = "DimAttributeLedgerFund_PSN";
                                break;

                            case "itemgroup":
                            case "itemgroups":
                                relatedDataEntity = "DimAttributeInventItemGroupEntity";
                                break;

                            case "item":
                            case "items":
                                relatedDataEntity = "DimAttributeInventTableEntity";
                                break;

                            case "job":
                            case "jobs":
                                relatedDataEntity = "DimAttributeHcmJobEntity";
                                break;

                            case "legalentity":
                            case "legalentities":
                                relatedDataEntity = "DimAttributeCompanyInfoEntity";
                                break;
                            case "posregister":
                            case "posregisters":
                                relatedDataEntity = "DimAttributeRetailTerminalEntity";
                                break;

                            case "position":
                            case "positions":
                                relatedDataEntity = "DimAttributeHcmPositionEntity";
                                break;

                            case "projectcontract":
                            case "projectcontracts":
                                relatedDataEntity = "DimAttributeProjInvoiceTableEntity";
                                break;

                            case "projectgroup":
                            case "projectgroups":
                                relatedDataEntity = "DimAttributeProjGroupEntity";
                                break;

                            case "project":
                            case "projects":
                                relatedDataEntity = "DimAttributeProjTableEntity";
                                break;

                            case "prospect":
                            case "prospects":
                                relatedDataEntity = "DimAttributeSmmBusRelTableEntity";
                                break;

                            case "resourcegroup":
                            case "resourcegroups":
                                relatedDataEntity = "DimAttributeWrkCtrResourceGroupEntity";
                                break;

                            case "resource":
                            case "resources":
                                relatedDataEntity = "DimAttributeWrkCtrTableEntity";
                                break;

                            case "store":
                            case "stores":
                                relatedDataEntity = "DimAttributeRetailStoreEntity";
                                break;

                            case "valuestream":
                            case "valuestreams":
                                relatedDataEntity = "DimAttributeOMValueStreamEntity";
                                break;

                            case "vendorgroup":
                            case "vendorgroups":
                                relatedDataEntity = "DimAttributeVendGroupEntity";
                                break;

                            case "vendor":
                            case "vendors":
                                relatedDataEntity = "DimAttributeVendTableEntity";
                                break;

                            case "worker":
                            case "workers":
                                relatedDataEntity = "DimAttributeHcmWorkerEntity";
                                break;

                            default:
                                relatedDataEntity = "DimAttributeFinancialTagEntity";
                                break;
                        }
                        #endregion
                        
                        var relations = dataEntityExtension.Relations;

                        if (relations.Where<Metadata.MetaModel.AxDataEntityViewRelation>(x => x.RelatedDataEntity == relatedDataEntity).Count<Metadata.MetaModel.AxDataEntityViewRelation>() == 0)
                        {
                            Metadata.MetaModel.AxDataEntityViewRelation dataEntityViewRelation = new Metadata.MetaModel.AxDataEntityViewRelation();
                            Metadata.MetaModel.AxDataEntityViewRelationConstraintField dataEntityViewRelationConstraintField = new Metadata.MetaModel.AxDataEntityViewRelationConstraintField();

                            dataEntityViewRelation.RelatedDataEntity = relatedDataEntity;
                            dataEntityViewRelation.Validate = Metadata.Core.MetaModel.NoYes.No;
                            dataEntityViewRelation.Cardinality = Metadata.Core.MetaModel.Cardinality.ZeroMore;
                            dataEntityViewRelation.Name = field.Name;
                            dataEntityViewRelation.RelatedDataEntityCardinality = Metadata.Core.MetaModel.RelatedTableCardinality.ZeroOne;
                            dataEntityViewRelation.RelatedDataEntityRole = $"{Guid.NewGuid().ToString("N")}";
                            dataEntityViewRelation.Role = $"{Guid.NewGuid().ToString("N")}";
                            
                            dataEntityViewRelationConstraintField.Name = $"{field.Name}Constraint";
                            dataEntityViewRelationConstraintField.Field = field.Name;
                            dataEntityViewRelationConstraintField.RelatedField = "Value";

                            dataEntityViewRelation.Constraints.Add(dataEntityViewRelationConstraintField);
                            dataEntityExtension.Relations.Add(dataEntityViewRelation);

                            updated = true;
                        }
                    }

                    if (updated)
                    {
                        Metadata.MetaModel.ModelSaveInfo modelSaveInfo = new Metadata.MetaModel.ModelSaveInfo(this.MetadataProvider.DataEntityViewExtensions.GetModelInfo(extensionName).FirstOrDefault());
                        this.MetadataProvider.DataEntityViewExtensions.Update(dataEntityExtension, modelSaveInfo);

                        updatedDataEntities += $"* {extensionName}\n";
                    }
                }

                if (entities.Count == 0)
                {
                    throw new Exception($"No entity extension has been found in {modelName} model for DimensionCombinationEntity and/or DimensionSetEntity. Did you execute \"Add financial dimensions for OData in advance?\" ");
                }

                if (updatedDataEntities != string.Empty)
                {
                    CoreUtility.DisplayInfo($"Following data entities extensions have been updated:\n{updatedDataEntities}");
                }
            }
            catch (Exception ex)
            {
                CoreUtility.HandleExceptionWithErrorMessage(ex);
            }
        }
        #endregion
    }
}
