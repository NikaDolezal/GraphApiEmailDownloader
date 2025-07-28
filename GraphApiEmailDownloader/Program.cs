using Azure.Identity;
using log4net;
using log4net.Config;
using Microsoft.Graph;
using Microsoft.Graph.Admin.ServiceAnnouncement.Messages;
using Microsoft.Kiota.Abstractions.Extensions;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Threading.Tasks;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.ListView;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.TreeView;

namespace GraphApiEmailDownloader
{
    internal class Program
    {
        // logger
        private static readonly ILog _logger = LogManager.GetLogger(typeof(Program));
        static async Task Main()
        {
            var logRepository = LogManager.GetRepository(Assembly.GetEntryAssembly());
            XmlConfigurator.Configure(logRepository, new FileInfo("log4net.config"));
            
            _logger.Info("---PRG START---");

            //-------------------------------------------------------------------------------------
            // load settings from cfg - no other input method supported
            Settings settings = Settings.LoadSettings();
            _logger.Info("OK - load settings");
            _logger.Debug("client ID: " + settings.ClientId);
            _logger.Debug("tenant ID: " + settings.TenantId);
            _logger.Debug("email: " + settings.Email);
            _logger.Debug("start date: " + settings.StartDate);
            _logger.Debug("dir path: " + settings.DirPath);

            //-------------------------------------------------------------------------------------
            // user auth through browser UI
            InitializeGraph( settings );
            _logger.Info("OK - initialize Graph");


            //-------------------------------------------------------------------------------------
            // process inbox of given email from given date - loaded from cfg
            await GraphUtil.ProcessInboxAsync( settings );
            _logger.Info("OK - process inbox");

            //-------------------------------------------------------------------------------------
            // save new start date in cfg
            DateTime timestamp = DateTime.UtcNow;
            settings.StartDate = String.Format( "{0}T{1}Z", timestamp.ToString( "yyyy-MM-dd" ), timestamp.ToString( "HH:mm:ss" ) );
            Settings.SaveSettings( settings );
            _logger.Info("OK - update cfg");

            _logger.Info("---PRG FINISH--");
        }

        static void InitializeGraph( Settings settings )
        {
            GraphUtil.InitializeGraphForUserAuth( settings,
                ( info, cancel ) =>
                {
                    // display device code message to user
                    Console.WriteLine( info.Message );
                    return Task.FromResult( 0 );
                } );
        }      

    }
}
