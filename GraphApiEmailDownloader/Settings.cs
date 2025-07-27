using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
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

        //-----------------------------------------------------------------------------------------
        //load settings: return type Settings filled from section 'settings' in specified .json
        public static Settings LoadSettings()
        {
            // specify cfg file
            IConfiguration config = new ConfigurationBuilder()
                // cfg 'appsettings.json' is required
                .AddJsonFile("appsettings.json", optional: false)
                .Build();
            // specify section and return settings            
            return config.GetRequiredSection("Settings").Get<Settings>() ??
                throw new Exception("Failed to load settings." );
        }

    }
}
