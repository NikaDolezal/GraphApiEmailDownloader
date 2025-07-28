using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Threading.Tasks;

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
            return config.Get<Settings>() ??
                throw new Exception( "Failed to load settings. Please refer to README." );
        }
        public static void SaveSettings( Settings settings )
        {

            //-------------------------------------------------------------------------------------
            // prepare new settings
            var newAppSettings = JsonSerializer.Serialize( settings, new JsonSerializerOptions() { WriteIndented = true } );

            //-------------------------------------------------------------------------------------
            // write cfg file
            // 'using' - for cleanup
            using ( StreamWriter outputFile = new StreamWriter( Path.Combine( AppDomain.CurrentDomain.BaseDirectory, "appsettings.json") ) )
            {
                outputFile.WriteLine(newAppSettings);
            }            
        }
    }
}
