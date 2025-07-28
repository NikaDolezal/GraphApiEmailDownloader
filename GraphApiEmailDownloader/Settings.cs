using log4net;
using Microsoft.Extensions.Configuration;
using System;
using System.IO;
using System.Text.Json;

namespace GraphApiEmailDownloader
{
    internal class Settings
    {
        public string ClientId { get; set; }
        public string TenantId { get; set; }
        public string[] GraphUserScopes { get; set; }
        public string Email { get; set; }
        public string StartDate { get; set; }
        public string DirPath { get; set; }
        //logger
        private static readonly ILog _logger = LogManager.GetLogger( typeof( Settings ) );


        //-----------------------------------------------------------------------------------------
        //load settings: return type Settings filled from section 'settings' in specified .json
        public static Settings LoadSettings()
        {
            // specify cfg file
            IConfiguration config = new ConfigurationBuilder()
                //set path to base = where .exe is
                .SetBasePath( AppDomain.CurrentDomain.BaseDirectory )
                // cfg 'appsettings.json' is required
                .AddJsonFile( "appsettings.json", optional: false )
                .Build();
            // return settings
            Settings settings = config.Get<Settings>();
            if( settings == null ) {
                throw new Exception("Failed to load settings. Please refer to README.");
            }
            return settings;
                
        }
        public static void SaveSettings( Settings settings )
        {

            //-------------------------------------------------------------------------------------
            // prepare new settings
            var newAppSettings = JsonSerializer.Serialize( settings, new JsonSerializerOptions() { WriteIndented = true } );

            //-------------------------------------------------------------------------------------
            // write cfg file
            // 'using' - for cleanup
            try
            {
                using (StreamWriter outputFile = new StreamWriter(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "appsettings.json")))
                {
                    outputFile.WriteLine(newAppSettings);
                }
            }
            catch( IOException e)
            {
                _logger.Error(String.Format("Problem writing cfg.\nFullpath: {0}\nError message: {1}", AppDomain.CurrentDomain.BaseDirectory, e.Message));

            }

        }
    }
}
