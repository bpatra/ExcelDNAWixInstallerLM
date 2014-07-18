using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Deployment.WindowsInstaller;

namespace mySampleWiXSetupCA
{
    /// <summary>
    /// Wraps the Session of custom action for logging.
    /// </summary>
    public class SessionLogger : ILogger
    {
        private readonly Session _session;
        public SessionLogger(Session session)
        {
            if(session == null) throw new ArgumentNullException(paramName: "session");
            _session = session;
        }
        public void Log(string message)
        {
            _session.Log(message);
        }
    }
}
