using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.SharePoint;

namespace BLL
{
    public class tabProcedury
    {
        const string targetList = "Procedury"; //"tabProcedury";

        public static bool AddNew(SPWeb web, string nazwaProcedury)
        {
            bool result = false;

            SPList list = web.Lists.TryGetList(targetList);

            //if (list!=null)
            //{
            SPListItem item = list.AddItem();
            item["Title"] = nazwaProcedury;

            try
            {
                item.SystemUpdate();
            }
            catch (Exception)
            { }
            finally
            {
                result = true;
            }
            //}

            return result;
        }

        public static int GetID(SPWeb web, string nazwaProcedury, bool createIfNotExist)
        {
            int result = 0;

            SPList list = web.Lists.TryGetList(targetList);

            //if (list != null)
            //{
            SPListItem item = list.Items.Cast<SPListItem>()
                .Where(i => i.Title == nazwaProcedury)
                .ToList()
                .FirstOrDefault();

            if (item != null)
            {
                result = item.ID;
            }
            //}

            if (result == 0 && createIfNotExist)
            {
                try
                {
                    item = list.AddItem();
                    item["Title"] = nazwaProcedury;
                    item.SystemUpdate();
                }
                catch (Exception)
                { }
                finally
                {
                    item = list.Items.Cast<SPListItem>()
                       .Where(i => i.Title == nazwaProcedury)
                       .ToList()
                       .FirstOrDefault();

                    result = item.ID;
                }
            }

            return result;
        }


        /// <summary>
        /// Na podstawie wprowadzonego tematu zadania sprawdza czy istnieje taka procedura, jeżeli nie istnieje to ją dodaje w statusie nowy,
        /// jeżeli istnieje i jest zatwierdzona to zwraca jej ID. W pozostałych przypadkach zwraca 0;
        /// </summary>
        /// <param name="web"></param>
        /// <param name="procName"></param>
        /// <returns></returns>
        public static int Ensure(SPWeb web, string procName)
        {
            SPList list = web.Lists.TryGetList(targetList);

            SPListItem item = list.Items.Cast<SPListItem>()
                .Where(i => i.Title == procName)
                .FirstOrDefault();

            if (item == null)
            {
                SPListItem newItem = list.AddItem();
                newItem["Title"] = procName;
                newItem["enumStatusProcedury"] = "Nowa";
                newItem.SystemUpdate();
                return newItem.ID;
            }
            else
            {
                return item.ID;
            }


        }

        public static int Get_TerminRealizacjiOfsetById(SPWeb web, int procId)
        {
            SPList list = web.Lists.TryGetList(targetList);
            //if (list!=null)
            //{
            SPListItem item = list.GetItemById(procId);
            if (item != null)
            {
                return item["colOczekiwanyTerminRealizacji_Ofset"] != null ? Int16.Parse(item["colOczekiwanyTerminRealizacji_Ofset"].ToString()) : 0;
            }

            //}

            return 0;
        }

        public static int Get_OperatorById(SPWeb web, int procId)
        {
            SPList list = web.Lists.TryGetList(targetList);
            //if (list != null)
            //{
            SPListItem item = list.GetItemById(procId);
            if (item != null)
            {
                return item["selDedykowanyOperator"] != null ? new SPFieldLookupValue(item["selDedykowanyOperator"].ToString()).LookupId : 0;
            }

            //}

            return 0;
        }
    }
}
