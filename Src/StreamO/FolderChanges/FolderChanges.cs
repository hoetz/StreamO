using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace StreamO.FolderChanges
{
    /// <summary>
    /// Provides facilities for receiving changes that occured in users' folders.
    /// </summary>
    public class FolderChanges
    {
        private ExchangeCredentials exchangeCredentials;
        private ExchangeVersion exchangeVersion = ExchangeVersion.Exchange2010_SP1;

        /// <summary>
        /// Initializes a new instance with specified Exchange credentials. Defaults to Exchange2010_SP1 Server version.
        /// </summary>
        /// <param name="exchangeCredentials">Credentials with permission to impersonate relevant user mailboxes</param>
        public FolderChanges(ExchangeCredentials exchangeCredentials)
        {
            this.exchangeCredentials = exchangeCredentials; 
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="exchangeCredentials"></param>
        /// <param name="exchangeVersion"></param>
        public FolderChanges(ExchangeCredentials exchangeCredentials, ExchangeVersion exchangeVersion)
            :this(exchangeCredentials)
        {
            this.exchangeVersion = exchangeVersion;
        }

        /// <summary>
        /// AutoDiscovers the EWS Url for the desired user and delivers all changes for the supplied folders / sync states
        /// </summary>
        /// <param name="userMailAddress">The owner of the targeted folders</param>
        /// <param name="folderStates">The targeted folders and their sync states</param>
        /// <returns></returns>
        public IEnumerable<SyncChangesResult> GetChangesFor(string userMailAddress, IEnumerable<FolderSyncState> folderStates)
        {
            var service = CreateImpersonatedService(new ImpersonatedUserId(ConnectingIdType.SmtpAddress, userMailAddress), this.exchangeCredentials);
            service.AutodiscoverUrl(userMailAddress,x=>true);

            foreach (var fld in folderStates)
            {
                List<ItemChange> fldChanges = new List<ItemChange>();
                string currentFldState = fld.SyncState;

                bool moreChangesAvailable;
                do
                {
                    var changes = service.SyncFolderItems(fld.TargetFolderId, PropertySet.IdOnly, null, 512, SyncFolderItemsScope.NormalItems, currentFldState);
                    fldChanges.AddRange(changes);
                    currentFldState = changes.SyncState;

                    // If more changes are available, issue additional SyncFolderItems requests.
                    moreChangesAvailable = changes.MoreChangesAvailable;
                }
                while (moreChangesAvailable);

                yield return new SyncChangesResult(new FolderSyncState(currentFldState,fld.TargetFolderId),fldChanges);
            }
        }


        private ExchangeService CreateImpersonatedService(ImpersonatedUserId UserId,ExchangeCredentials exchangeCredentials)
        {
            return new ExchangeService(this.exchangeVersion) { Credentials = exchangeCredentials, ImpersonatedUserId=UserId };
        }
    }
}
