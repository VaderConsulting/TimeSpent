//
// Copyright © Microsoft Corporation.  All Rights Reserved.
// This code released under the terms of the 
// Microsoft Public License (MS-PL, http://opensource.org/licenses/ms-pl.html)
//
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

namespace TimeSpent
{
    class Program
    {
        static void Main(string[] args)
        {
            CmdLine cmd = new CmdLine();
            if (cmd.Parse(args))
                cmd.Run();
        }
    }
}
