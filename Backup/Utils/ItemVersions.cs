
using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;


namespace Backup
{
    public class ItemVersions
    {
        public string createdDateTime { get; set; }
        public int ID { get; set; }


        public List<Dictionary<string,object>> versions { get; set; }
        public Dictionary<string,object> currentItems { get; set; }


        public ItemVersions(string createdDateTime, int ID)
        {
            this.createdDateTime = createdDateTime;
            this.ID = ID;
            versions = new List<Dictionary<string,object>>();
            currentItems = new Dictionary<string,object>();
            
        }

    }



}
