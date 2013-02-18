//------------------------------------------------------------------------------
// wgsgrep.cs
//------------------------------------------------------------------------------
// Simple command line interface for Windows (MSN) Desktop Search
// <Revision: 2007-02-03>
//
// Required runtime environment:
//  - Windows Desktop Search
//  - Microsoft .NET Framework 2.0 or higher
// Required software to build this program:
//  - Visual Studio 2005 (Free edition is available)
//
// Copyright (C) 2005-2007 INOUE Tomohiro <ml <at> noue.org>
// All rights reserved.
//
// Redistribution and use in source and binary forms, with or without
// modification, are permitted provided that the following conditions
// are met:
// 1. Redistributions of source code must retain the above copyright
//    notice, this list of conditions and the following disclaimer.
// 2. Redistributions in binary form must reproduce the above copyright
//    notice, this list of conditions and the following disclaimer in the
//    documentation and/or other materials provided with the distribution.
//
// THIS SOFTWARE IS PROVIDED BY THE TEAM AND CONTRIBUTORS ``AS IS'' AND
// ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE
// IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR
// PURPOSE ARE DISCLAIMED.  IN NO EVENT SHALL THE TEAM OR CONTRIBUTORS BE
// LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR
// CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF
// SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR
// BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY,
// WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE
// OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN
// IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
//------------------------------------------------------------------------------
//
// 
//
//------------------------------------------------------------------------------

using System;
using System.Data;
using System.Data.OleDb;
using System.Collections;
using System.Text.RegularExpressions;
using System.Collections.Generic;
using Microsoft.Win32;
using Microsoft.Search.Interop;

namespace wdsgrep
{
	/// <summary>
	/// Sample C# App using the QueryBuilder sample wrapper.  
	/// </summary>
	public class WDSgrep
	{
		private const string revision = "4.0";
        private Regex urlregex = null;
        private bool out_slash = false;
        private const string ext_default = "mew";
        private const string ifilter_src = "eml";
        //string regist_default = ext_default + "=" + ifilter_src; // "mew=eml"

        private int wds_version = 0;
        private string wds_verstr = "";
        private string sortcols = null;
        private string sortrank = null;
        //private string selectcol = null;

#if DEBUG
        private bool DEBUG = true;
#else
        private bool DEBUG = false;
#endif
        //private Microsoft.Windows.DesktopSearch.Query._Recordset myRecordSet;			//ADO RecordSet
        //private Microsoft.Windows.DesktopSearch.Query.SearchDesktopClass mySDObj;		//WDS Query interface

		public WDSgrep(string[] args)
		{
            Console.Out.NewLine = "\n";

            string regist_target = null;
            string path = null;
            string ext  = null;
            string fullregexp = null;
            string query = "";
            string query_orig = "";
            //string type = null;
            bool mailonly = false;
            
            bool path_fragment = false;
            //bool fallback = true;

            versioncheck(); // exit if < 3.0 

            // parse args
            if (args.Length <= 0)
            {
                usage(); // and exit
            }

            // xxx  "-X" makes COM error on WDS 2.6
            IEnumerator e = args.GetEnumerator();
            while (e.MoveNext())
	        {
                switch ((string)e.Current)
                {
                    /*
                    case "-q":
                        break;  // just ignore for compatibility to gdsgrep
                    */
                    case "-h":
                        usage(); // and exit
                        break; 
                    case "-R":
                        if (e.MoveNext()) regist_target = (string)e.Current;
                        register(regist_target); // and exit
                        break;
                    case "-p":
                        if (! e.MoveNext()) usage();
                        path = ((string)e.Current).Replace('\\', '/');
                        break;
                    case "-p2": // compatible to rev. 2.5
                        if (!e.MoveNext()) usage();
                        path = ((string)e.Current).Replace('\\', '/');
                        path_fragment = true;
                        break;
                    case "-e":
                        if (!e.MoveNext()) usage();
                        ext = (string)e.Current;
                        if (ext == "") usage();
                        ext = ext_withoutdot(ext);
                        break;
                    case "-E":
                        if (!e.MoveNext()) usage();
                        fullregexp = (string)e.Current;
                        break;
                    /*
                    case "-P":
                        if (!e.MoveNext()) usage();
                        type = (string)e.Current;
                        if (type == "") usage();
                        break;
                     */
                    case "-S":
                        if (!e.MoveNext()) usage();
                        sortcols = (string)e.Current;
                        if (sortcols == "") usage();
                        break;
                    case "-Sr":
                        sortcols = sortrank;
                        break;
                    case "-m": // for compatibility to gdsgrep and mew
                        mailonly = true;
                        break;
                    case "-s":
                        out_slash = true;
                        break;
                    case "--debug":
                        DEBUG = true;
                        break;
                    default: // other args
                        query_orig += " " + e.Current;
                        break;
                }
                
	        }
            if (query_orig == "") usage(); // and exit
            query = query_orig;

            if (mailonly) ext = ext_default; // mew
            //type = TypeFilter.Email;// "communications/e-mail"

            // Creat Regex to restrict results in file extension and path
            string regexstr = null;
            if (path != null) regexstr += "^" + Regex.Escape(path) + ".*";
            if (ext != null)
            {
                if (regexstr == null) regexstr = ".*";
                regexstr += Regex.Escape(ext_withdot(ext)) + "$";
            }
            
            if (fullregexp != null) regexstr = fullregexp;
            if (regexstr != null) urlregex = new Regex(regexstr, RegexOptions.IgnoreCase);

            //if (ext != null) query += " " + ext_withoutdot(ext); // "mew" is faster than ".mew"
            if (ext != null) query += " ext:" + ext_withoutdot(ext); // a little fanster than "mew"

            string query_no_path = query;
            // if path_frargment 
            if (path != null)
            {
                if (path_fragment) // default: false
                {   // add path fragments to query words
                    // remove {MyDocuments} since words extracted from it don't hit
                    // todo: other Shell special folders might involve same problems.
                    System.Globalization.CultureInfo invariant = System.Globalization.CultureInfo.InvariantCulture;

                    string mydocument = System.Environment.GetFolderPath(Environment.SpecialFolder.Personal);
                    mydocument = mydocument.ToLower(invariant);
                    string path2 = String.Copy(path); // for safety for future incompatibility
                    path2 = path2.Replace('/', '\\').ToLower(invariant);
                    if (path2.StartsWith(mydocument)) path2 = path2.Remove(0, mydocument.Length); //ignore case

                    Regex drivereg = new Regex("[a-z]:", RegexOptions.IgnoreCase);
                    //string[] dirs = path2.Split(new char[] { '/', '\\' });
                    string[] dirs = path2.Split(new char[] { '\\' });
                    foreach (string frag in dirs)
                    {
                        if (!drivereg.IsMatch(frag)) // ignore drive letter
                            query += " " + frag;
                    }
                }
                else
                {
                    string folder;
                    if (wds_version <= 2) // legacy code
                    { // "c\path\foo"
                        //folder = String.Join("", path.Split(new char[] {':'})).Replace('/', '\\');
                        //folder = '"' + path.Remove(path.IndexOf(':'), 1).Replace('/', '\\') + '"';
                        folder = '"' + path.Replace(":", "").Replace('/', '\\') + '"';
                    }
                    else // WDS 3.0
                    { // c:\path\foo (even if a path includes white spaces), "c:\path\foo" does not match sub folders
                        folder = path.Replace('/', '\\').TrimEnd(new char[] { '\\' });
                    }
                    query += " folder:" + folder;
                }
            }

            if (! WDSHelper.IsDesktopSearchInstalled)
            {
                Console.Error.WriteLine("Error: WDS (ver. 3.x or above) cannot detected.");
                Environment.Exit(-1);
            }

            int hit = run(query, urlregex);
            /*
            if (hit == 0 && path != null)
            {
                query = query_no_path + " folder:" + path.Replace('/', '\\');
                hit = run(query, type, urlregex); //
            }
             */

            if (DEBUG)
            {
                Console.WriteLine("WDS version: {0} ({1})", wds_version, wds_verstr);
                Console.WriteLine("Query:\t" + query);
                Console.WriteLine("Hits on WDS: " + hit); // todo xxx
                Console.WriteLine("Path:\t"  + path);
                Console.WriteLine("Regex:\t" + urlregex);
            }
#if DEBUG
                Console.Read(); //for debug
#endif
        }

		[STAThread]
		static void Main(string[] args) 
		{
            new WDSgrep(args);
        }

        private void usage()
        {
            string usage =
                "wdsgrep " + revision + ": A simple command line interface for Windows Desktop Search\n" + 
                "usage: wdsgrep [options] <query words>\r\n" +
                "options: -p <path>    Results are limited in the absolute path\n" +
                "         -e <ext>     Results are limited to files with the extention '.ext'\n" +
                "         -E <regexp>  Limit to regular expression (overrides -p and -e)\n" +
//                "         -P <type>    Set PerceivedTypes filter (see MS SDK documents)\n" +
                "         -S <columns> Sort by the column. Add DESC for reversed output\n" +
                "                      A comma separated list of columns is also available\n" +
                "                      Default: sort by date (-S '" + sortcols + "')\n" +
                "         -Sr          Sort by rank (-S '"+ sortrank + "')\n" +
                "         -m           Mew mode (equivalent to '-e mew')\n" +
                "         -s           Use slash instead of backslash for path delimiter\n" +
//                "         -q           (ignored for compatibility)\n" +
                "         -R [dst=src] Register '.dst' to WDS as the same file type of '.src'\n" +
                "                      Default: mew=eml; register .mew files as .eml (e-mails)\n";
            Console.Write(usage);
            Environment.Exit(-1);
        }

        private void versioncheck()
        {
            const string WDSreg = "SOFTWARE\\Microsoft\\Windows Desktop Search";

            RegistryKey rk = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(WDSreg);
            if (rk == null)
            {
                string key = Microsoft.Win32.Registry.LocalMachine.Name + "\\" + WDSreg;
                Console.Error.WriteLine("Failed to access registry: " + key);
                Console.Error.WriteLine("WDS is not installed?");
                Environment.Exit(-1);
            }
            object ver = rk.GetValue("Version");
            if (ver == null)
            {
                Console.Error.WriteLine("Failed to read version: " + rk.Name + "\\Version");
                rk.Close();
                Environment.Exit(-1);
            }
            string verstr = (string)ver;
            int majorver = int.Parse(verstr.Substring(0, verstr.IndexOf('.')));
            wds_verstr = verstr; 
            wds_version = majorver;

            if (majorver <= 2) { // WDS ver. 2
                Console.Error.WriteLine("Installed WDS version: {0} ({1})", wds_version, wds_verstr);
                Console.Error.WriteLine("This program requires WDS 3.0 or higer");
                Environment.Exit(-1);
                //sortcols  = "PrimaryDate DESC";
                //sortrank  = "Rank DESC";
                //selectcol = "Url";
            }
            else { // WDS ver 3
                sortcols  = "System.ItemDate DESC";
                sortrank  = "System.Search.Rank DESC";
                //selectcol = "System.ItemUrl";
            }

#if DEBUG
		    Console.WriteLine("WDS version: {0} ({1})", wds_version, wds_verstr);
  
#endif
            return;
        }

        private string ext_withdot(string ext)
        {
            string rtn = ext;
            if (ext[0] != '.') rtn = "." + ext;
            return rtn;
        }

        private string ext_withoutdot(string ext)
        {
            string rtn = ext;
            if (ext[0] == '.') rtn = ext.Substring(1,ext.Length-1);
            return rtn;
        }

        private void register(string target) // target: "[.]dst=[.]src"
        {
            string dst = ext_default; // .mew
            string src = ifilter_src; // .eml

            if (target != null)
            {
                String[] dst_src = target.Split(new Char[] { '=' }, 2);
                if (dst_src.Length != 2) usage(); // and exit
                dst = dst_src[0];
                src = dst_src[1];
                if (dst.Substring(0, 1) != ".") dst = "." + dst;
                if (src.Substring(0, 1) != ".") src = "." + src;
            }
            if (DEBUG) Console.WriteLine("Registering '{0}' using the value of '{1}'.", dst, src);

            try
            {
                object src_val = null;
                string regpath;
                RegistryKey parent, src_key, dst_key;
                if (wds_version <= 2)
                {
                    regpath = "Software\\Microsoft\\RSSearch\\ContentIndexCommon\\Filters\\Extension";
                    parent = Registry.CurrentUser.OpenSubKey(regpath, true);
                    if (parent == null)
                    {
                        string key = Registry.CurrentUser.Name + "\\" + regpath;
                        Console.Error.WriteLine("Failed to access registry: " + key);
                        Environment.Exit(-1);
                    }
                    src_key = parent.OpenSubKey(src);
                }
                else {
                    regpath = "SOFTWARE\\Classes";
                    //string srcpath = regpath + "\\" + src;
                    parent = Registry.LocalMachine.OpenSubKey(regpath, true);
                    src_key = parent.OpenSubKey(src);
                }
                if (src_key == null) {
                    string key = parent.Name + "\\" + src;
                    Console.Error.WriteLine("Failed to access registry: " + key);
                    Console.Error.WriteLine("src '{0}' cannot be found.", src);
                    parent.Close();
                    Environment.Exit(-1);
                }
                // rksrc.GetValueKind is not available for .NET 1.1 
                src_val = src_key.GetValue(null);
                if (src_val == null) {
                    Console.Error.WriteLine("Failed to read registry: " + src_key.Name);
                    Console.Error.WriteLine("Unexpected format.", src);
                    parent.Close(); src_key.Close();
                    Environment.Exit(-1);
                }
                //if (DEBUG) Console.WriteLine("The value is '{0}'.", src_val.ToString());

                dst_key = parent.OpenSubKey(dst, true);
                if (dst_key == null) dst_key = parent.CreateSubKey(dst);
                src_key.Close();
                parent.Close();

                bool backup = true;
                object dst_val = dst_key.GetValue(null);
                if (dst_val == null)
                {
                    backup = false;
                }
                else{
                    Type type = dst_val.GetType();
                    if (type == src_val.GetType())
                    {
                        if (type == typeof(string))
                        {
                            if ((string)dst_val == (string)src_val) backup = false;
                        }
                        else if (type == typeof(string[]))
                        {
                            string[] da = (string[])dst_val;
                            string[] sa = (string[])src_val;
                            if (da.Length == sa.Length)
                            {
                                backup = false;
                                for (int i = 0; i < da.Length; i++)
                                {
                                    if (da[i] != sa[i]) backup = false;
                                }
                            }
                        }
                    }
                }
                if (backup)
                {
                    string backupkey = "wdsgrep-backup";
                    dst_key.SetValue(backupkey, dst_val); 
                    Console.WriteLine("Previous value is copied to {0} {1}", dst_key.Name, backupkey);
                }
                dst_key.SetValue(null, src_val);
                dst_key.Close();
                Console.WriteLine("Registration succeeded.");
                Console.WriteLine("'{0}' is registered to WDS as the same file type of '{1}'.", dst, src);
#if DEBUG
                Console.Read(); //for debug
#endif
            }
            catch (Exception)
            {
                Console.Error.WriteLine("Failed to access registry.");
                Environment.Exit(-1);
                //throw;
            }
            Environment.Exit(0);
        }

        private int run(string query, Regex urlregex)
        {
            //myQuery.setSortColumnList(sortcols);

            // Search
            List<String> results = null;
            try { 
                results = WDSHelper.ExecuteQuery(query, sortcols);
            }
            catch (System.Data.OleDb.OleDbException e)
            {
                //if (DEBUG) Console.Error.WriteLine(e.ToString());
                Console.Error.WriteLine(e.ToString());
            }
            catch (System.Runtime.InteropServices.COMException e)
            {
                //if (DEBUG) Console.Error.WriteLine(e.ToString());
                Console.Error.WriteLine(e.ToString());
            }
            if (results == null) return -1;

            // Show Results with regex filter
            int count = 0;
            foreach (String url in results)
            {
                string path = null;
                if (url.StartsWith("file:"))  // may always be true 
                    path = url.Substring(5, url.Length - 5);
                // restrict path and file extention
                if (urlregex == null || urlregex.IsMatch(path))
                {
                    if (!out_slash) path = path.Replace("/", "\\");
                    Console.WriteLine(path);
                    count++;
                }
            }
            return count;
        }

	}

    class WDSHelper
    {
        // Shared connection used for any search index queries
        private static OleDbConnection conn;

        /// <summary>
        /// Initialized connection.  Silently fails (check IsDesktopSearchInstalled property)
        /// </summary>
        static WDSHelper()
        {
            try
            {
                conn = new OleDbConnection("Provider=Search.CollatorDSO;Extended Properties='Application=Windows';");
                conn.Open();
            }
            catch
            {
                conn = null;
            }
        }

        /// <summary>
        /// Indicates if WDS 3.0 is installed based on initial connection succeeding or not
        /// </summary>
        public static bool IsDesktopSearchInstalled
        {
            get
            {
                return (conn != null);
            }
        }

        /// <summary>
        /// Performs search for a given term and returns results as a list of strings
        /// </summary>
        /// <param name="data">Search term</param>
        /// <returns>Search results</returns>
        public static List<string> ExecuteQuery(string aqsQuery, string sortcols)
        {
            //string selectcol = "System.ItemURL";

            CSearchManager manager = new CSearchManagerClass();
            CSearchCatalogManager catalogManager = manager.GetCatalog("SystemIndex");
            CSearchQueryHelper queryHelper = catalogManager.GetQueryHelper();
            queryHelper.QuerySorting = sortcols;
            string sqlQuery = queryHelper.GenerateSQLFromUserQuery(aqsQuery);  

            OleDbCommand cmd = new OleDbCommand();
            cmd.Connection = conn;
            cmd.CommandText = sqlQuery;
            /*
            cmd.CommandText = string.Format(
                "SELECT {0} FROM systemindex WHERE CONTAINS(*, '{1}')", // search contents and properties
                // "SELECT {0} FROM systemindex WHERE CONTAINS('{1}')", // search contents only
                selectcol, aqsQuery);
             */
            OleDbDataReader results = cmd.ExecuteReader();

            List<string> items = new List<string>();
            while (results.Read())
            {
                items.Add(results.GetString(0));
            }

            results.Close();

            return items;
        }
    }

}

