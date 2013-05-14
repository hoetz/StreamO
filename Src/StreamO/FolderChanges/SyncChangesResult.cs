using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Exchange.WebServices.Data;

namespace StreamO.FolderChanges
{
    /// <summary>
    /// Contains information about the changes that occured in a Folder.
    /// </summary>
    public class SyncChangesResult
    {
        private FolderSyncState _FolderSyncState;

        private IList<ItemChange> _ItemChanges = new List<ItemChange>();
        public IEnumerable<ItemChange> ItemChanges
        {
            get { return this._ItemChanges; }
        }

        public FolderId FolderId
        {
            get { return _FolderSyncState.TargetFolderId; }
        }

        /// <summary>
        /// The current sync state of the folder. May be stored for later sync attempts.
        /// </summary>
        public string CurrentSyncState
        {
            get { return _FolderSyncState.SyncState; }
        }

        public SyncChangesResult(FolderSyncState folderSyncState, IEnumerable<ItemChange> ItemChanges)
        {
            this._FolderSyncState = folderSyncState;
            if (folderSyncState == null)
                throw new ArgumentException("folderSyncState");

            foreach (var item in ItemChanges)
            {
                this._ItemChanges.Add(item);
            }
        }

    }
}
