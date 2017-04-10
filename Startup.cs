using Microsoft.Owin;
using Owin;

[assembly: OwinStartupAttribute(typeof(LaborNeedsSchedulingFramework.Startup))]
namespace LaborNeedsSchedulingFramework
{
    public partial class Startup
    {
        public void Configuration(IAppBuilder app)
        {
            ConfigureAuth(app);
        }
    }
}
