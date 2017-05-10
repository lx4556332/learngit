//
// import System.dll
// import Castle.Windsor.dll
// import PengeSoft.Core.dll
// import PengeSoft.Enterprise.dll
// import PengeSoft.EntEx.dll
// import PengeSoft.Service.dll
// import PS.Core.Commons.dll
// import PS.Core.Config.dll
// import PS.Core.PsReport.dll
// import FIXF.Core.Services.VotingPlatform.dll
// import FIXF.Core.Services.Common.dll
// import FIXF.Core.Services.Finance.dll
//

using System;
using Castle.Windsor;
using FIXF.VotingPlatform.Service.Vote;
using FIXF.Core.Services.VotingPlatform;
using PengeSoft.EntEx;
using PengeSoft.WorkZoneData;
using PS.Core.Commons;
using PS.Core.Config;
using PS.Core.PsReport;
using System.Collections.Generic;
using System.Text.RegularExpressions;

public class DataSvr : IPsReportDataSvr
{
    public StatisticsItem GetReportData(string token, WindsorContainer iocContainer, User currUser, ReportDefine define,
        ReportParam param)
    {
        StatisticsItem result = new StatisticsItem();
        try
        {
            string topicvoteid = param.GetParamValue<string>("docCode");
            string webName = ConfigManager.GetSysSettingByString("PublicWebName"); //公众服务网站名称
            string webUrl = ConfigManager.GetSysSettingByString("PublicWebAccessUrl"); //公众服务平台访问地址
            string tel = ConfigManager.GetSysSettingByString("PublicTel"); //公众服务咨询电话
            string publicAddress = ConfigManager.GetSysSettingByString("PublicAddress"); //公众服务办事大厅地址

            IVoteMgeSvr svr = (IVoteMgeSvr) iocContainer["FIXF.VoteMgeSvr"];

            ParticipantQueryPara queryParam = new ParticipantQueryPara();
            queryParam.SetTopicVoteId(topicvoteid); //投票主题ID
            ParticipantList participantList = svr.GetParticipantList(token, queryParam, 0, -1, false); //参与投票

            VoteTopic topic = svr.GetVoteTopic(token, topicvoteid, true, string.Empty); //表决

            result.AddProperty("DocCode", DateTime.Now.ToString("yyyyMMddHHmm"));
            result.AddProperty("EaAreaName", topic.EaAreaName);
            result.AddProperty("PrintDate", DateTime.Now.ToString("D"));
            result.AddProperty("VoteTitle", topic.VoteTitle);
            result.AddProperty("RangeIntro", topic.RangeIntro);
            result.AddProperty("BeginTime", "投票通知公告时间：从"+topic.BeginTime.ToString("D") + "起");
            result.AddProperty("EndTime", "投票表决时间：从" + topic.BeginTime.ToString("D") + "至" + topic.EndTime.ToString("D") + "止 ");
            result.AddProperty("DetailIntro",
                topic.DetailIntro.Length > 0 ? TextNoHTML(topic.DetailIntro.Replace("&nbsp;", "")) : "");

            result.AddProperty("PublicTel", tel);
            result.AddProperty("PublicAddress", publicAddress);
            result.AddProperty("PublicWebAccessUrl", webUrl);
            result.AddProperty("PublicWebName", webName);

            StatisticsList subkeyList = new StatisticsList();

            if (topic != null)
            {
                Dictionary<string, string> builsDic = new Dictionary<string, string>(); //楼栋、单元
                Dictionary<string, VoteObject> housesDic = new Dictionary<string, VoteObject>(); //房屋

                PSLog4net.Info(this, "楼栋-单元拼装");
                VoteObjectList voteObjects = topic.VoteObjects;
                foreach (VoteObject item in topic.VoteObjects)
                {
                    if (!IsVote(participantList, item))
                    {
                        housesDic.Add(item.KeyId, item);
                        string key = item.TopicVoteId + item.EaBuildName + item.EaUnitName;
                        string val = item.EaBuildName + item.EaUnitName;

                        if (!builsDic.ContainsKey(key))
                        {
                            builsDic.Add(key, val);
                        }
                    }
                }
                //-------------------------
                SubKeyList subKeyLists = topic.SubKeys;
                foreach (SubKey subKeyItem in subKeyLists) //表决事项
                {
                    PSLog4net.Info(this, "表决事项拼装");
                    string canItems = "";
                    string subItems = "";
                    canItems += subKeyItem.ItemTitle + "；";
                    subItems += "" + Enum.GetName(typeof(ECandidateType), subKeyItem.CandidateType) + "：";
                    foreach (CandidateItem canItem in subKeyItem.CandidateItems)
                    {
                        subItems += canItem.CandiNumber + "：" + canItem.CandiTitle + " ";
                    }
                    //-----------------------
                    StatisticsItem subKey = new StatisticsItem();
                    subKey.AddProperty("subKeys", "<div style='margin-left:50px'>" + canItems + "</div>");
                    subKey.AddProperty("subItems", "<div style='margin-left:50px'>" + subItems + "</div>");
                    subKey.AddProperty("subKeySdescribe",
                            "最少选择：" + subKeyItem.MinChoice + "项，最多选择：" + subKeyItem.MaxChoice + "项");
                  

                    StatisticsList groupList = new StatisticsList();
                    foreach (KeyValuePair<string, string> buildItem in builsDic) //楼栋-单元
                    {
                        StatisticsItem group = new StatisticsItem();
                        group.AddProperty("buildCode", buildItem.Key);
                        group.AddProperty("buildName", buildItem.Value);

                        StatisticsList dataList = new StatisticsList();
                        int cnt = 1;
                        foreach (string key in housesDic.Keys) //房屋
                        {

                            VoteObject voteObject = housesDic[key];
                            if (buildItem.Value == (voteObject.EaBuildName + voteObject.EaUnitName))
                            {
                                StatisticsItem dataItem = dataList.AddNew();

                                dataItem.AddProperty("Index", cnt);
                                dataItem.AddProperty("EaHouseDimension", voteObject.Dimensions);
                                dataItem.AddProperty("Room", voteObject.Room);
                                dataItem.AddProperty("Floor", voteObject.Floor);
                                cnt++;
                            }

                        }

                        group.AddProperty("dataList", dataList);
                        groupList.Add(group);
                    }
                    subKey.AddProperty("groupList", groupList);
                    subkeyList.Add(subKey);
                }
            }

            result.AddProperty("subkeyList", subkeyList);
        }
        catch (Exception e)
        {
            PSLog4net.Error(this, e.Message);
        }

        return result;
    }

    /// <summary>
    /// 是否已经投票
    /// </summary>
    /// <returns></returns>
    public bool IsVote(ParticipantList participantList, VoteObject voteObject)
    {
        foreach (Participant participant in participantList)
        {
            if (participant.VoteObjectId == voteObject.KeyId && !string.IsNullOrEmpty(participant.CandidateItemNos))
            {
                return true;
            }
        }
        return false;
    }

    public string TextNoHTML(string Htmlstring)
    {
        if (Htmlstring.Length > 0)
        {
            //删除脚本  
            Htmlstring = Regex.Replace(Htmlstring, @"<script[^>]*?>.*?</script>", "", RegexOptions.IgnoreCase);
            //删除HTML  
            Htmlstring = Regex.Replace(Htmlstring, @"<(.[^>]*)>", "", RegexOptions.IgnoreCase);
            Htmlstring = Regex.Replace(Htmlstring, @"([/r/n])[/s]+", "", RegexOptions.IgnoreCase);
            Htmlstring = Regex.Replace(Htmlstring, @"-->", "", RegexOptions.IgnoreCase);
            Htmlstring = Regex.Replace(Htmlstring, @"<!--.*", "", RegexOptions.IgnoreCase);
            Htmlstring = Regex.Replace(Htmlstring, @"&(quot|#34);", "/", RegexOptions.IgnoreCase);
            Htmlstring = Regex.Replace(Htmlstring, @"&(amp|#38);", "&", RegexOptions.IgnoreCase);
            Htmlstring = Regex.Replace(Htmlstring, @"&(lt|#60);", "<", RegexOptions.IgnoreCase);
            Htmlstring = Regex.Replace(Htmlstring, @"&(gt|#62);", ">", RegexOptions.IgnoreCase);
            Htmlstring = Regex.Replace(Htmlstring, @"&(nbsp|#160);", "   ", RegexOptions.IgnoreCase);
            Htmlstring = Regex.Replace(Htmlstring, @"&(iexcl|#161);", "/xa1", RegexOptions.IgnoreCase);
            Htmlstring = Regex.Replace(Htmlstring, @"&(cent|#162);", "/xa2", RegexOptions.IgnoreCase);
            Htmlstring = Regex.Replace(Htmlstring, @"&(pound|#163);", "/xa3", RegexOptions.IgnoreCase);
            Htmlstring = Regex.Replace(Htmlstring, @"&(copy|#169);", "/xa9", RegexOptions.IgnoreCase);
            Htmlstring = Regex.Replace(Htmlstring, @"&#(/d+);", "", RegexOptions.IgnoreCase);
            //替换掉 < 和 > 标记
            Htmlstring.Replace("<", "");
            Htmlstring.Replace(">", "");
            Htmlstring.Replace("/r/n", "");
            //返回去掉html标记的字符串
        }
        return Htmlstring;
    }
}
