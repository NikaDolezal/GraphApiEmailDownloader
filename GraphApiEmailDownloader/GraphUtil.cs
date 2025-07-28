using Azure.Identity;
using log4net;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

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
        // logger
        private static readonly ILog _logger = LogManager.GetLogger(typeof(GraphUtil));


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
        
        public static async Task<MessageCollectionResponse> ProcessInboxAsync( Settings settings )                  
        {           
            // make sure client isn't null
            if (_userClient == null)
            {
                //to deal with one layer up
                throw new System.NullReferenceException( "Graph not initialized" );
            }

            //-------------------------------------------------------------------------------------
            // first request
            // user from cfg email
            var res = await _userClient.Users[settings.Email]
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
                    config.QueryParameters.Select = new[] { "receivedDateTime", "subject", "body" };
                    // filter by date received
                    string dateFlt = String.Format( "receivedDateTime ge {0}", settings.StartDate );
                    _logger.Debug( "receivedDateTime flt: " + dateFlt );
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
                    return SaveMessage( message, settings.DirPath );
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

        private static bool SaveMessage(Message message, string dirPath)
        {
            //-------------------------------------------------------------------------------------
            // create target fullpath
            string tgtDir = @dirPath;

            string msgDir;
            // check subject
            if (!String.IsNullOrEmpty(message.Subject))
            {
                // remove illegal chars
                msgDir = string.Join("", message.Subject.Split(Path.GetInvalidFileNameChars()));
            }
            else
            {
                // arbitrarily name with 'DateTimeReceived' to handle multiple no subject messages in same dir
                string tmp = String.Format( "no subject {0}", message.ReceivedDateTime );
                // 'DateTimeReceived' contains illegal chars
                msgDir = string.Join("", tmp.Split(Path.GetInvalidFileNameChars()));
            }
                        
            // create directory for specific message
            string tgtFullpath = Path.Combine(tgtDir, msgDir);
            Directory.CreateDirectory( tgtFullpath );
            _logger.Debug("target fullpath: " + tgtFullpath);

            //-------------------------------------------------------------------------------------
            // write file
            // 'using' - for cleanup
            try
            {
                using (StreamWriter outputFile = new StreamWriter(Path.Combine(tgtFullpath, "message.txt")))
                {
                    outputFile.WriteLine(message.Body.Content);
                }
            }
            catch (IOException e)
            {
                _logger.Error( String.Format( "Problem writing message to file.\nFullpath: {0}\nError message: {1}", tgtFullpath, e.Message ) );
            }

            return true;
        }
    }
}
