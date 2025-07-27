using Azure;
using Azure.Core;
using Azure.Identity;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Kiota.Abstractions;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using static System.Windows.Forms.LinkLabel;

namespace GraphApiEmailDownloader
{
    internal class GraphUtil
    {
        // settings from cfg
        private static Settings _settings;
        // user auth token credential
        private static DeviceCodeCredential _deviceCodeCredential;  
        // graph client from user auth
        public static GraphServiceClient _userClient;

        public static void InitializeGraphForUserAuth( Settings settings,
            Func<DeviceCodeInfo, CancellationToken, Task> deviceCodePrompt )
        {
            _settings = settings;

            // options from cfg settings
            var options = new DeviceCodeCredentialOptions
            {
                ClientId = _settings.ClientId,
                TenantId = _settings.TenantId,
                DeviceCodeCallback = deviceCodePrompt,
            };

            _deviceCodeCredential = new DeviceCodeCredential( options );

            _userClient = new GraphServiceClient( _deviceCodeCredential, settings.GraphUserScopes );
        }
        
        public static async Task<MessageCollectionResponse> ProcessInboxAsync( string email, string startDate )                  
        {           
            // make sure client isn't null
            if (_userClient == null)
            {
                //Console.WriteLine( "Graph not initialized" );
                //return null;
                throw new System.NullReferenceException( "Graph not initialized" );
            }

            //-------------------------------------------------------------------------------------
            // first request
            // user from cfg email
            var res = await _userClient.Users[email]
                // only 'Inbox'
                .MailFolders["Inbox"]
                .Messages
                .GetAsync( ( config ) =>
                {
                    // content as text
                    config.Headers.Add( "Prefer", "outlook.body-content-type='text'" );
                    // testing PageIterator
                    config.QueryParameters.Top = 3;
                    // request specific properties
                    config.QueryParameters.Select = new[] { "from", "receivedDateTime", "subject", "body" };
                    // filter by date received
                    string dateFlt = String.Format( "receivedDateTime ge {0}", startDate );
                    Console.WriteLine( "date flt: {0}",dateFlt );
                    config.QueryParameters.Filter = dateFlt;                   
                    // sort by date - from newest
                    config.QueryParameters.Orderby = new[] { "receivedDateTime DESC" };                    
                } );
            //-------------------------------------------------------------------------------------
            // PageIterator
            var pageIterator = PageIterator<Message, MessageCollectionResponse>.CreatePageIterator( 
                _userClient,
                res,
                // callback for each message
                ( message ) =>
                {
                    // TODO - proof
                    return SaveMessage( message );
                },
                // configure subsequent page requests
                ( req ) =>
                {
                    // add header
                    req.Headers.Add( "Prefer", "outlook.body-content-type=\"text\"" );
                    return req;
                } );

            await pageIterator.IterateAsync();

            return res;
        }
        
        private static bool SaveMessage( Message message )
        {
            //-------------------------------------------------------------------------------------
            // create target fullpath
            // TODO - load from cfg
            string tgtDir = "C:\\Users\\root\\Desktop\\test";
            
            // remove illegal chars
            string msgDir = string.Join("", message.Subject.Split( Path.GetInvalidFileNameChars() ) );
            
            //TODO - empty subject
            // create directory for specific message
            string tgtFullpath = Path.Combine( tgtDir, msgDir );
            Directory.CreateDirectory( tgtFullpath );

            //-------------------------------------------------------------------------------------
            // write file
            // 'using' - for cleanup
            using ( StreamWriter outputFile = new StreamWriter( Path.Combine( tgtFullpath, "message.txt" ) ) )
            {
                outputFile.WriteLine( message.Body.Content );
            }

            // TODO - proof
            return true;
        }
    }
}
