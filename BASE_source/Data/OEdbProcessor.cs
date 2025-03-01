using DocumentFormat.OpenXml.Bibliography;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace BASE.Data
{
    public class OEdbProcessors
    {
        // functions
        public static List<string> GetParticipants(string names)
        {
            // Split by comma and trim spaces
            List<string> participantsList = new List<string>(names.Split(','));

            // Trim each name to remove extra spaces
            for (int i = 0; i < participantsList.Count; i++)
            {
                participantsList[i] = participantsList[i].Trim();
            }
            return participantsList;
        }


    }
    public class PracticeArea_Element // 1 Prac_Group
    {
        public string AcronymName { get; set; } // eg CAR, DAR
        public List<Practice_Element> practice_Elements { get; set; } = new List<Practice_Element>();
    }

    public class Practice_Element // 2 Prac_OU
    {
        public string CodeAndNumber { get; set; }
        public string CharStr { get; set; }
        public List<string> WeaknessesList { get; set; } = new List<string>();
        public List<string> StrengthList { get; set; } = new List<string>();
        public List<string> RecommendationList { get; set; } = new List<string>();
        public List<string> SessionList { get; set; } = new List<string>();
        public List<string> ParticipantList { get; set; } = new List<string>();
        public List<PrjSup_Element> PrjSup_Elements { get; set; } = new List<PrjSup_Element>();

        public void UpdateParticipants(List<string> participantsToAdd)
        {
            ParticipantList = ParticipantList.Union(participantsToAdd).ToList();
        }

        public void UpdateSessions(string sessionNameToAdd)
        {
            if (string.IsNullOrEmpty(sessionNameToAdd)) return;
            sessionNameToAdd = sessionNameToAdd?.Trim();
            if (!SessionList.Contains(sessionNameToAdd))
            {
                SessionList.Add(sessionNameToAdd);
            }
            SessionList.Sort();
        }

        public void UpdateWeaknesses(string weaknessToAdd)
        {
            if (string.IsNullOrEmpty(weaknessToAdd)) return;
            weaknessToAdd = weaknessToAdd?.Trim();
            if (!WeaknessesList.Contains(weaknessToAdd))
            {
                WeaknessesList.Add(weaknessToAdd);
            }
            WeaknessesList.Sort();
        }

        public void UpdateStrengths(string strengthToAdd)
        {
            if (string.IsNullOrEmpty(strengthToAdd)) return;
            strengthToAdd = strengthToAdd?.Trim();
            if (!StrengthList.Contains(strengthToAdd))
            {
                StrengthList.Add(strengthToAdd);
            }
            StrengthList.Sort();
        }

        public void UpdateRecommendations(string recommendationToAdd)
        {
            if (string.IsNullOrEmpty(recommendationToAdd)) return;

            recommendationToAdd = recommendationToAdd?.Trim();
            if (!RecommendationList.Contains(recommendationToAdd))
            {
                RecommendationList.Add(recommendationToAdd);
            }
            RecommendationList.Sort();
        }


    }

    public class PrjSup_Element // 4 Prac_Instan
    {
        public string CodeAndNumber { get; set; }
        public string projectSupportName { get; set; }
        public List<string> participantList { get; set; } = new List<string>();
        // Public static method so it can be accessed from other classes
        public string sessionName { get; set; }
        public string weaknessStr { get; set; }
        public string strengthStr { get; set; }
        public string recommendationStr { get; set; }
        public string Char { get; set; } 
        public List<OE_Element> oE_Elements { get; set; } = new List<OE_Element>();


    }

    public enum E_OEStatus
    {
        OkFile,
        OkDirectory,
        NotOk,
        None
    }
    public enum E_YesNo
    {
        Yes,
        No
    }

    public class OE_Element // 5 OE
    {
        public string ProjectName { get; set; }
        public E_OEStatus OeStatus { get; set; } = E_OEStatus.None; // ok file
        public E_YesNo Sufficient { get; set; } = E_YesNo.No; // default 
    }


}
