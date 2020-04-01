using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Configuration;

namespace JINL
{
    class Program
    {
        static void Main(string[] args)
        {
            ClientContext psCtx = LoginPnPCore();

            //SpCsPnpcoreCreatePropertyBag(psCtx);
            //SpCsPnpcoreReadPropertyBag(psCtx);
            //SpCsPnpcorePropertyBagExists(psCtx);
            //SpCsPnpcorePropertyBagIndex(psCtx);
            //SpCsPnpcoreDeletePropertyBag(psCtx);

            Console.WriteLine("Done");
            Console.ReadLine();
        }

        //gavdcodebegin 01
        static void SpCsPnpcoreCreatePropertyBag(ClientContext spCtx)
        {
            List myList = spCtx.Web.Lists.GetByTitle("TestList");

            myList.SetPropertyBagValue("myKey", "myValueString");
        }
        //gavdcodeend 01

        //gavdcodebegin 02
        static void SpCsPnpcoreReadPropertyBag(ClientContext spCtx)
        {
            List myList = spCtx.Web.Lists.GetByTitle("TestList");

            string myKeyValue = myList.GetPropertyBagValueString("myKey", "");
            Console.WriteLine(myKeyValue);
        }
        //gavdcodeend 02

        //gavdcodebegin 03
        static void SpCsPnpcorePropertyBagExists(ClientContext spCtx)
        {
            List myList = spCtx.Web.Lists.GetByTitle("TestList");

            bool myKeyExists = myList.PropertyBagContainsKey("myKey");
            Console.WriteLine(myKeyExists.ToString());
        }
        //gavdcodeend 03

        //gavdcodebegin 04
        static void SpCsPnpcorePropertyBagIndex(ClientContext spCtx)
        {
            List myList = spCtx.Web.Lists.GetByTitle("TestList");

            myList.AddIndexedPropertyBagKey("myKey");
            IEnumerable<string> myIndexedPropertyBagKeys = 
                                                myList.GetIndexedPropertyBagKeys();

            foreach (string oneKey in myIndexedPropertyBagKeys)
            {
                Console.WriteLine(oneKey);
            }
        }
        //gavdcodeend 04

        //gavdcodebegin 05
        static void SpCsPnpcoreDeletePropertyBag(ClientContext spCtx)
        {
            List myList = spCtx.Web.Lists.GetByTitle("TestList");

            myList.RemovePropertyBagValue("myKey");
        }
        //gavdcodeend 05

        //-------------------------------------------------------------------------------
        static ClientContext LoginPnPCore()
        {
            OfficeDevPnP.Core.AuthenticationManager pnpAuthMang =
                new OfficeDevPnP.Core.AuthenticationManager();
            ClientContext rtnContext =
                        pnpAuthMang.GetSharePointOnlineAuthenticatedContextTenant
                            (ConfigurationManager.AppSettings["spUrl"],
                             ConfigurationManager.AppSettings["spUserName"],
                             ConfigurationManager.AppSettings["spUserPw"]);

            return rtnContext;
        }
    }
}
