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

        public static void MergeStringIntoList(string aString, ref List<string> list)
        {
            if (string.IsNullOrEmpty(aString)) return;
            aString = aString?.Trim();
            if (!list.Contains(aString))
            {
                list.Add(aString);
            }
            list.Sort();
        }

        public static void MergeListIntoList(List<string> listToaAdd, ref List<string> listToUpdate)
        {
            listToUpdate = listToUpdate.Union(listToaAdd).ToList();
        }

    }
    public class PracticeArea_Element // 1 Prac_Group
    {
        public string AcronymName { get; set; } // eg CAR, DAR
        public List<Practice_Element> practice_Elements { get; set; } = new List<Practice_Element>();
        public override string ToString()
        {
            return $"{AcronymName} ({practice_Elements.Count})";
        }
    }

    public class Practice_Element // 2 Prac_OU
    {
        public string CodeAndNumber { get; set; }
        public string Char { get; set; }
        public List<string> WeaknessesList { get; set; } = new List<string>();
        public List<string> StrengthList { get; set; } = new List<string>();
        public List<string> RecommendationList { get; set; } = new List<string>();
        public List<string> SessionList { get; set; } = new List<string>();
        public List<string> ParticipantList { get; set; } = new List<string>();
        public List<Process_Element> ProcessElements { get; set; } = new List<Process_Element>();
        public List<PrjSup_Element> PrjSup_Elements { get; set; } = new List<PrjSup_Element>();

        public override string ToString()
        {
            return $"{CodeAndNumber} ({ProcessElements.Count},{PrjSup_Elements.Count})";
        }

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
        public void UpdateSessions(List<string> sessionListIn)
        {
            SessionList = SessionList.Union(sessionListIn).ToList();
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
        public void UpdateWeaknesses(List<string> weaknessesToAdd)
        {
            WeaknessesList = WeaknessesList.Union(weaknessesToAdd).ToList();
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

        public void UpdateStrengths(List<string> strenthsToAdd)
        {
            StrengthList = StrengthList.Union(strenthsToAdd).ToList();
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

        public void UpdateRecommendations(List<string> recommendationsToAdd)
        {
            RecommendationList = RecommendationList.Union(recommendationsToAdd).ToList();   
        }

    }

    public class Process_Element // 3 Process
    {
        public string CodeAndNumber { get; set; }
        public string ProcessName { get; set; }
        public string Char { get; set; }
        public List<string> WeaknessesList { get; set; } = new List<string>();
        public List<string> StrengthList { get; set; } = new List<string>();
        public List<string> RecommendationList { get; set; } = new List<string>();
        public List<string> SessionList { get; set; } = new List<string>();
        public List<string> ParticipantList { get; set; } = new List<string>();
        public List<PrjSup_Element> PrjSup_Elements { get; set; } = new List<PrjSup_Element>();

        public override string ToString()
        {
            return $"{CodeAndNumber}:{ProcessName} ({PrjSup_Elements.Count})";
        }

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
        public void UpdateSessions(List<string> sessionListIn)
        {
            SessionList = SessionList.Union(sessionListIn).ToList();
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
        public void UpdateWeaknesses(List<string> weaknessesToAdd)
        {
            WeaknessesList = WeaknessesList.Union(weaknessesToAdd).ToList();
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

        public void UpdateStrengths(List<string> strenthsToAdd)
        {
            StrengthList = StrengthList.Union(strenthsToAdd).ToList();
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

        public void UpdateRecommendations(List<string> recommendationsToAdd)
        {
            RecommendationList = RecommendationList.Union(recommendationsToAdd).ToList();
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
        public override string ToString()
        {
            return $"{CodeAndNumber}:{projectSupportName} ({oE_Elements.Count})";
        }

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
        public override string ToString()
        {
            return $"{ProjectName}";
        }
    }


}
