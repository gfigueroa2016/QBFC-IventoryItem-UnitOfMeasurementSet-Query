using System;
using Interop.QBFC13;

namespace QBFC_IventoryItem_UnitOfMeasurementSet_Query
{
    class Program
    {
        public static void Main()
        {
            QBSessionManager sessionManager = new QBSessionManager();
            try
            {
                IMsgSetRequest requestMsgSet = sessionManager.CreateMsgSetRequest("US", 13, 0);
                sessionManager.OpenConnection("", "TEST");
                sessionManager.BeginSession("", ENOpenMode.omDontCare);
                requestMsgSet.Attributes.OnError = ENRqOnError.roeContinue;
                IItemInventoryQuery ItemInventoryQueryRq = requestMsgSet.AppendItemInventoryQueryRq();
                IMsgSetResponse responseMsgSet = sessionManager.DoRequests(requestMsgSet);
                if (responseMsgSet == null) return;
                IResponseList responseList = responseMsgSet.ResponseList;
                if (responseList == null) return;
                for (int i = 0; i < responseList.Count; i++)
                {
                    IResponse response = responseList.GetAt(i);
                    if (response.StatusCode >= 0)
                    {
                        if (response.Detail != null)
                        {
                            ENResponseType responseType = (ENResponseType)response.Type.GetValue();
                            if (responseType == ENResponseType.rtItemInventoryQueryRs)
                            {
                                IItemInventoryRetList ItemInventoryRet = (IItemInventoryRetList)response.Detail;
                                for (int j = 0; j < ItemInventoryRet.Count; j++)
                                {
                                    if (ItemInventoryRet.GetAt(j).UnitOfMeasureSetRef != null)
                                    {
                                        string itemListId = ItemInventoryRet.GetAt(j).ListID.GetValue();
                                        string itemFullName = ItemInventoryRet.GetAt(j).FullName.GetValue();
                                        short ItemType = ItemInventoryRet.GetAt(j).Type.GetValue();
                                        string unitOfMeasurementRefListId = ItemInventoryRet.GetAt(j).UnitOfMeasureSetRef.ListID.GetValue();
                                        string unitOfMeasurementRefFullName = ItemInventoryRet.GetAt(j).UnitOfMeasureSetRef.FullName.GetValue();
                                        short unitOfMeasurementRefType = ItemInventoryRet.GetAt(j).UnitOfMeasureSetRef.Type.GetValue();
                                        Console.WriteLine($"Item Id: {itemListId}\r\nItem Full Name: {itemFullName}\r\n{ItemType}\r\nUnit of Measurement List Id: {unitOfMeasurementRefListId}\r\nUnit of Measurement Full Name: {unitOfMeasurementRefFullName}\r\nUnit of Measurement Type: {unitOfMeasurementRefType}\r\n");
                                        IMsgSetRequest uomRequestMsgSet = sessionManager.CreateMsgSetRequest("US", 13, 0);
                                        IUnitOfMeasureSetQuery UnitOfMeasureSetQueryRq = uomRequestMsgSet.AppendUnitOfMeasureSetQueryRq();
                                        UnitOfMeasureSetQueryRq.ORListQuery.ListIDList.Add(ItemInventoryRet.GetAt(j).UnitOfMeasureSetRef.ListID.GetValue());
                                        IMsgSetResponse unitOfMeaurementResponseMsgSet = sessionManager.DoRequests(uomRequestMsgSet);
                                        WalkUnitOfMeasureSetQueryRs(unitOfMeaurementResponseMsgSet);
                                    }
                                }
                            }
                        }
                    }
                }
                sessionManager.EndSession();
                sessionManager.CloseConnection();
                Console.ReadLine();
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message, "Error");
                sessionManager.EndSession();
                sessionManager.CloseConnection();
                Console.ReadLine();
            }
        }
        static void WalkUnitOfMeasureSetQueryRs(IMsgSetResponse responseMsgSet)
        {
            if (responseMsgSet == null) return;
            IResponseList responseList = responseMsgSet.ResponseList;
            if (responseList == null) return;
            for (int i = 0; i < responseList.Count; i++)
            {
                IResponse response = responseList.GetAt(i);
                if (response.StatusCode >= 0)
                {
                    if (response.Detail != null)
                    {
                        ENResponseType responseType = (ENResponseType)response.Type.GetValue();
                        if (responseType == ENResponseType.rtUnitOfMeasureSetQueryRs)
                        {
                            IUnitOfMeasureSetRetList UnitOfMeasureSetRet = (IUnitOfMeasureSetRetList)response.Detail;
                            for (int j = 0; j < UnitOfMeasureSetRet.Count; j++)
                            {
                                if (UnitOfMeasureSetRet != null)
                                {
                                    if (UnitOfMeasureSetRet.GetAt(j).BaseUnit != null)
                                    {
                                        if (UnitOfMeasureSetRet.GetAt(j).BaseUnit.Abbreviation != null)
                                        {
                                            string baseUnitAbbreviation = UnitOfMeasureSetRet.GetAt(j).BaseUnit.Abbreviation.GetValue();
                                            Console.WriteLine($"Base Unit Abbreviation: {baseUnitAbbreviation}");
                                        }
                                        if (UnitOfMeasureSetRet.GetAt(j).BaseUnit.Name != null)
                                        {
                                            string baseUnitName = UnitOfMeasureSetRet.GetAt(j).BaseUnit.Name.GetValue();
                                            Console.WriteLine($"Base Unit Name: {baseUnitName}");
                                        }
                                        if (UnitOfMeasureSetRet.GetAt(j).BaseUnit.Type != null)
                                        {
                                            short baseUnitType = UnitOfMeasureSetRet.GetAt(j).BaseUnit.Type.GetValue();
                                            Console.WriteLine($"Base Unit Type: {baseUnitType}");
                                        }
                                    }
                                    if (UnitOfMeasureSetRet.GetAt(j).DefaultUnitList != null)
                                    {
                                        for (int i22940 = 0; i22940 < UnitOfMeasureSetRet.GetAt(j).DefaultUnitList.Count; i22940++)
                                        {
                                            if (UnitOfMeasureSetRet.GetAt(j).DefaultUnitList.GetAt(i22940) != null)
                                            {
                                                IDefaultUnit DefaultUnit = UnitOfMeasureSetRet.GetAt(j).DefaultUnitList.GetAt(i22940);
                                                if (DefaultUnit.UnitUsedFor != null)
                                                {
                                                    ENUnitUsedFor UnitUsedFor22941 = DefaultUnit.UnitUsedFor.GetValue();
                                                    Console.WriteLine($"Unit used for: {UnitUsedFor22941}");
                                                }
                                                if (DefaultUnit.Unit != null)
                                                {
                                                    string Unit22942 = DefaultUnit.Unit.GetValue();
                                                    Console.WriteLine($"Default Unit: {Unit22942}");
                                                }
                                            }
                                        }
                                    }
                                    if (UnitOfMeasureSetRet.GetAt(j).EditSequence != null)
                                    {
                                        string editSequence = UnitOfMeasureSetRet.GetAt(j).EditSequence.GetValue();
                                        Console.WriteLine($"Edit sequence: {editSequence}");
                                    }
                                    if (UnitOfMeasureSetRet.GetAt(j).IsActive != null)
                                    {
                                        bool isActive = UnitOfMeasureSetRet.GetAt(j).IsActive.GetValue();
                                        Console.WriteLine($"Is active: {isActive}");
                                    }
                                    if (UnitOfMeasureSetRet.GetAt(j).ListID != null)
                                    {
                                        string listId = UnitOfMeasureSetRet.GetAt(j).ListID.GetValue();
                                        Console.WriteLine($"List Id: {listId}");
                                    }
                                    if (UnitOfMeasureSetRet.GetAt(j).Name != null)
                                    {
                                        string name = UnitOfMeasureSetRet.GetAt(j).Name.GetValue();
                                        Console.WriteLine($"Name: {name}");
                                    }
                                    if (UnitOfMeasureSetRet.GetAt(j).RelatedUnitList != null)
                                    {
                                        for (int i22936 = 0; i22936 < UnitOfMeasureSetRet.GetAt(j).RelatedUnitList.Count; i22936++)
                                        {
                                            if (UnitOfMeasureSetRet.GetAt(j).RelatedUnitList.GetAt(i22936) != null)
                                            {
                                                IRelatedUnit RelatedUnit = UnitOfMeasureSetRet.GetAt(j).RelatedUnitList.GetAt(i22936);
                                                if (RelatedUnit.Name != null)
                                                {
                                                    string Name22937 = RelatedUnit.Name.GetValue();
                                                    Console.WriteLine($"Related Unit: {Name22937}");
                                                }
                                                if (RelatedUnit.Abbreviation != null)
                                                {
                                                    string Abbreviation22938 = RelatedUnit.Abbreviation.GetValue();
                                                    Console.WriteLine($"Related Unit Abbreviation: {Abbreviation22938}");
                                                }
                                                if (RelatedUnit.ConversionRatio != null)
                                                {
                                                    double ConversionRatio22939 = RelatedUnit.ConversionRatio.GetValue();
                                                    Console.WriteLine($"Conversion Ratio: {ConversionRatio22939}");
                                                }
                                            }
                                        }
                                    }
                                    if (UnitOfMeasureSetRet.GetAt(j).TimeCreated != null)
                                    {
                                        DateTime timeCreated = UnitOfMeasureSetRet.GetAt(j).TimeCreated.GetValue();
                                        Console.WriteLine($"Time created: {timeCreated}");
                                    }
                                    if (UnitOfMeasureSetRet.GetAt(j).TimeModified != null)
                                    {
                                        DateTime timeModified = UnitOfMeasureSetRet.GetAt(j).TimeModified.GetValue();
                                        Console.WriteLine($"Time modified: {timeModified}");
                                    }
                                    if (UnitOfMeasureSetRet.GetAt(j).Type != null)
                                    {
                                        short type = UnitOfMeasureSetRet.GetAt(j).Type.GetValue();
                                        Console.WriteLine($"Type: {type}");
                                    }
                                    if (UnitOfMeasureSetRet.GetAt(j).UnitOfMeasureType != null)
                                    {
                                        ENUnitOfMeasureType unitOfMeasurementType = UnitOfMeasureSetRet.GetAt(j).UnitOfMeasureType.GetValue();
                                        Console.WriteLine($"Unit of Measure Type: {unitOfMeasurementType}\r\n");
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
    }
}
