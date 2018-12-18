using Microsoft.Owin;
using Owin;

[assembly: OwinStartupAttribute(typeof(UploadExcelDemo.Startup))]
namespace UploadExcelDemo
{
    public partial class Startup
    {
        public void Configuration(IAppBuilder app)
        {
            ConfigureAuth(app);
        }
    }
}
