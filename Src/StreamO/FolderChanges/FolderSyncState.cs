using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Exchange.WebServices.Data;

namespace StreamO.FolderChanges
{
    /// <summary>
    /// Contains SyncState Information about an Exchange Folder
    /// </summary>
    public class FolderSyncState
    {
        private string _SyncState;
        public string SyncState
        {
            get { return this._SyncState; }
        }

        private FolderId _TargetFolderId;
        public FolderId TargetFolderId
        {
            get { return this._TargetFolderId; }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="syncState">The optional sync state representing the point in time when the folder was last synced. null indicates the earliest possible time.</param>
        /// <param name="folderId"></param>
        public FolderSyncState(string syncState, FolderId folderId)
        {
            this._SyncState = syncState;
            if (folderId == null)
                throw new ArgumentException("folderId");
            this._TargetFolderId = folderId;
        }

    }
}
