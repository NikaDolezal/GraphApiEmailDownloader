using Azure.Identity;
using Microsoft.Graph;
using Microsoft.Graph.Admin.ServiceAnnouncement.Messages;
using Microsoft.Kiota.Abstractions.Extensions;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Threading.Tasks;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.ListView;

namespace GraphApiEmailDownloader
{
    internal class Program
    {
        static async Task Main()
        {
            Console.WriteLine( "---PRG START---" );

            //-------------------------------------------------------------------------------------
            // load settings from cfg - no other input method supported
            Settings settings = Settings.LoadSettings();
            Console.WriteLine( "OK - load settings" );
            Console.WriteLine( "client ID: " + settings.ClientId );
            Console.WriteLine( "tenant ID: " + settings.TenantId );
            Console.WriteLine( "email: " + settings.Email );
            Console.WriteLine( "start date: " + settings.StartDate );
            Console.WriteLine( "dir path: " + settings.DirPath );

            //-------------------------------------------------------------------------------------
            // user auth through browser UI
            InitializeGraph( settings );
            Console.WriteLine( "OK - initialize Graph" );

            //-------------------------------------------------------------------------------------
            // process inbox of given email from given date - loaded from cfg
            await GraphUtil.ProcessInboxAsync( settings );
            Console.WriteLine( "OK - process inbox" );

            //-------------------------------------------------------------------------------------
            // save new start date in cfg
            DateTime timestamp = DateTime.UtcNow;
            settings.StartDate = String.Format( "{0}T{1}Z", timestamp.ToString( "yyyy-MM-dd" ), timestamp.ToString( "HH:mm:ss" ) );
            Settings.SaveSettings( settings );
            Console.WriteLine( "OK - update cfg" );

            Console.WriteLine( "---PRG FINISH--" );
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
