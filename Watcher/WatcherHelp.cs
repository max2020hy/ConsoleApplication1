using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.IO;

namespace Watcher
{
    public class WatcherHelp
    {
        public FileSystemWatcher Watcher { get; set; }
        public string DirPath { get; set; }
        /// <summary>
        /// 创建的新文件
        /// </summary>
        private string FilePath { get; set; }
        public WatcherHelp(string dirPath )
        {
            
            Watcher=new FileSystemWatcher();
            DirPath = dirPath;

        }
        public void Watching(Delegate watcherM)
        {
            Watcher.NotifyFilter = NotifyFilters.CreationTime;
            Watcher.Created += Watcher_Created;

        }

    
        private void Watcher_Created(object sender, FileSystemEventArgs e)
        {
          FilePath= e.FullPath;

        }
        public string GetChangeFile()
        {

            return FilePath;
        }
    }
}
