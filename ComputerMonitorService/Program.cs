using ComputerMonitorService;
using System.ServiceProcess;

public static class Program
{
    public static void Main()
    {
        ServiceBase[] ServicesToRun;
        ServicesToRun = new ServiceBase[]
        {
            new Service1()
        };
        ServiceBase.Run(ServicesToRun);
    }
}
