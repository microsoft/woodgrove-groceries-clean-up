using System;
using System.Threading.Tasks;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Extensions.Logging;

namespace Company.Function
{
    public class TimerCleanUpTrigger
    {
        private readonly ILogger _logger;

        public TimerCleanUpTrigger(ILoggerFactory loggerFactory)
        {
            _logger = loggerFactory.CreateLogger<TimerCleanUpTrigger>();
        }

        // https://learn.microsoft.com/en-us/azure/azure-functions/functions-bindings-timer
        [Function("TimerCleanUpTrigger")]
        public async Task Run([TimerTrigger("0 30 9 * * *")] TimerInfo myTimer) /*Occurs at 9:30 AM every day*/
        {
            _logger.LogInformation($"C# Timer trigger function executed at: {DateTime.Now}");
            
            if (myTimer.ScheduleStatus is not null)
            {
                _logger.LogInformation($"Next timer schedule at: {myTimer.ScheduleStatus.Next}");
            }

            GraphService graphService = new GraphService(_logger);
            await graphService.CleanUpDormantAccountsAsync();
        }
    }
}
