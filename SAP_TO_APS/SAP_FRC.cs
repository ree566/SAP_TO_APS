using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SAP.Middleware.Connector;

namespace SAP_TO_APS
{
    class SAP_FRC: IDestinationConfiguration
    {
        public RfcConfigParameters GetParameters(string destinationName)
        {
            if (destinationName.Equals("SAPMES"))
            {
                RfcConfigParameters rfcParams = new RfcConfigParameters();
                //rfcParams.Add(RfcConfigParameters.AppServerHost, "172.20.1.176");   //SAP主機IP
                rfcParams.Add(RfcConfigParameters.AppServerHost, "172.20.3.6");   //SAP主機IP
                //rfcParams.Add(RfcConfigParameters.SystemNumber, "05");              //SAP實例
                rfcParams.Add(RfcConfigParameters.SystemNumber, "06");              //SAP實例
                rfcParams.Add(RfcConfigParameters.Client, "168");                   // Client
                rfcParams.Add(RfcConfigParameters.User, "MES.ACL");                     //用戶名
                rfcParams.Add(RfcConfigParameters.Password, "MESMES");              //密碼
                rfcParams.Add(RfcConfigParameters.Language, "zf");                  //登陆語言
                //rfcParams.Add(RfcConfigParameters.PoolSize, "5");
                //rfcParams.Add(RfcConfigParameters.MaxPoolSize, "10");
                rfcParams.Add(RfcConfigParameters.ConnectionIdleTimeout, "1800");
                return rfcParams;
            }
            else
            {
                return null;
            }

        }

        public bool ChangeEventsSupported()
        {

            return false;

        }


        public event RfcDestinationManager.ConfigurationChangeHandler ConfigurationChanged;
    }
      

}
