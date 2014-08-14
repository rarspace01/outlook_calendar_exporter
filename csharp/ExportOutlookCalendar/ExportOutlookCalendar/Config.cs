using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DSF.src.Util
{
    class Config
    {

        public static string getTempPath()
        {
            string tempPath = System.IO.Path.GetTempPath() + ".calexport\\";
            return tempPath;
        }

    }
}
