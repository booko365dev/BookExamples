﻿using System;
using System.Configuration;
using System.Net;
using System.Security;
using Microsoft.SharePoint.Client;
using RestSharp;

namespace XDHG
{
    class Program
    {
        static void Main(string[] args)
        {
            //TestRestSharpGet();
            //TestRestSharepPost();

            Console.ReadLine();
        }

        //gavdcodebegin 001
        static RestClient LoginRestSharp() //*** LEGACY CODE *** 
        {
            //****************************************************************************
            // Microsoft removed the "SharePointOnlineCredentials" object from the
            // "Microsoft.SharePointOnline.CSOM" namespace (from version 16.1.20) to
            // prevent login with the credentials of a user. It is not possible to use
            // the code shown in this section any more
            //****************************************************************************

            var securePw = new SecureString();
            foreach (
                char oneChar in ConfigurationManager.AppSettings["spUserPw"].ToCharArray())
            {
                securePw.AppendChar(oneChar);
            }
            SharePointOnlineCredentials myCredentials = new SharePointOnlineCredentials(
                        ConfigurationManager.AppSettings["spUserName"], securePw);

            RestClient myClient = new RestClient(
                        ConfigurationManager.AppSettings["spUrl"] + "/_api/");
            myClient.CookieContainer = new CookieContainer();
            myClient.CookieContainer.SetCookies(new Uri(
                        ConfigurationManager.AppSettings["spUrl"]), 
                        myCredentials.GetAuthenticationCookie(new Uri(
                                    ConfigurationManager.AppSettings["spUrl"])));

            return myClient;
        }
        //gavdcodeend 001

        //gavdcodebegin 002
        static void TestRestSharpGet() //*** LEGACY CODE *** 
        {
            RestClient myClient = LoginRestSharp();

            RestRequest myRequestResult = new RestRequest(
                                "web/lists/getbytitle('TestList')/items", Method.GET);
            myRequestResult.AddHeader("Accept", "application/json");

            string resultJSON = myClient.Execute(myRequestResult).Content;
        }
        //gavdcodeend 002

        //gavdcodebegin 003
        static void TestRestSharepPost() //*** LEGACY CODE *** 
        {
            RestClient myClient = LoginRestSharp();

            RestRequest myRequestDigest = new RestRequest("contextinfo", Method.POST);
            myRequestDigest.AddHeader("Accept", "application/json");
            dynamic myDigest = myClient.Execute<dynamic>(myRequestDigest).Data;

            RestRequest myRequestResultC = RequestCreate(myDigest["FormDigestValue"]);
            RestRequest myRequestResultU = RequestUpdate(myDigest["FormDigestValue"]);
            RestRequest myRequestResultD = RequestDelete(myDigest["FormDigestValue"]);

            string resultJSONC = myClient.Execute(myRequestResultC).Content;
            string resultJSONU = myClient.Execute(myRequestResultU).Content;
            string resultJSOND = myClient.Execute(myRequestResultD).Content;
        }
        //gavdcodeend 003

        //gavdcodebegin 004
        static RestRequest RequestCreate(string Digest) //*** LEGACY CODE *** 
        {
            RestRequest myRequest = new RestRequest(
                    "web/lists/getbytitle('TestList')/items", Method.POST);
            myRequest.AddHeader("Accept", "application/json");
            myRequest.AddHeader("Content-Type", "application/json;odata=verbose");
            myRequest.AddHeader("X-RequestDigest", Digest);

            myRequest.AddParameter("application/json;odata=verbose", "{ '__metadata': " +
                    "{ 'type': 'SP.ListItem' }, 'Title': 'MyTestItem'}", 
                    ParameterType.RequestBody);

            return myRequest;
        }
        //gavdcodeend 004

        //gavdcodebegin 005
        static RestRequest RequestUpdate(string Digest) //*** LEGACY CODE *** 
        {
            RestRequest myRequest = new RestRequest(
                    "web/lists/getbytitle('TestList')/items(1)", Method.POST);
            myRequest.AddHeader("Accept", "application/json");
            myRequest.AddHeader("Content-Type", "application/json;odata=verbose");
            myRequest.AddHeader("X-RequestDigest", Digest);
            myRequest.AddHeader("IF-MATCH", "*");
            myRequest.AddHeader("X-HTTP-Method", "MERGE");

            myRequest.AddParameter("application/json;odata=verbose", "{ '__metadata': " +
                    "{ 'type': 'SP.ListItem' }, 'Title': 'MyItemUpdated'}", 
                    ParameterType.RequestBody);

            return myRequest;
        }
        //gavdcodeend 005

        //gavdcodebegin 006
        static RestRequest RequestDelete(string Digest) //*** LEGACY CODE *** 
        {
            RestRequest myRequest = new RestRequest(
                    "web/lists/getbytitle('TestList')/items(2)", Method.POST);
            myRequest.AddHeader("Accept", "application/json");
            myRequest.AddHeader("Content-Type", "application/json;odata=verbose");
            myRequest.AddHeader("X-RequestDigest", Digest);
            myRequest.AddHeader("IF-MATCH", "*");
            myRequest.AddHeader("X-HTTP-Method", "DELETE");

            return myRequest;
        }
        //gavdcodeend 006
    }
}
