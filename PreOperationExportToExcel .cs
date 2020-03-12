using Microsoft.Xrm.Sdk;
using System;


namespace MEA.Plugins
{
    public class PreOperationExportToExcel : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            IOrganizationServiceFactory factory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = factory.CreateOrganizationService(context.UserId);
            var parentContext = context.ParentContext;
           
            if (parentContext != null && (parentContext.MessageName == "ExportToExcel" || parentContext.MessageName == "ExportDynamicToExcel"))
            {
                Entity user = service.Retrieve("systemuser", context.UserId, new Microsoft.Xrm.Sdk.Query.ColumnSet("firstname", "lastname"));
                String userName = user.GetAttributeValue<string>("firstname") + " " + user.GetAttributeValue<string>("lastname");
                OrganizationRequest actionRequest = new OrganizationRequest("mea_ActionSendEmailOnExportToExcel");
                actionRequest["UserName"] = userName;
                actionRequest["Entity"] = context.PrimaryEntityName;
                OrganizationResponse response = service.Execute(actionRequest);
            }
            

        }
    }
}
