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

            //-------------------------------------------------------------------------------------
            // user auth through browser UI
            InitializeGraph( settings );
            Console.WriteLine( "OK - intialize Graph" );

            //-------------------------------------------------------------------------------------
            // process inbox of given email from given date - loaded from cfg
            await GraphUtil.ProcessInboxAsync( settings.Email, settings.StartDate );
            Console.WriteLine( "OK - process inbox" );

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
