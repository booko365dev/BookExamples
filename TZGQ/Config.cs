using Microsoft.TeamsFx.Configuration;

namespace TZGQ
{
    public class ConfigOptions
    {
        public TeamsFxOptions TeamsFx { get; set; }
    }
    public class TeamsFxOptions
    {
        public AuthenticationOptions Authentication { get; set; }
    }
}
