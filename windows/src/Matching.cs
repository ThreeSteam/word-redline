using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

static class MatchFiles
{
    public static string Kind(string path) {
        string s=Path.GetFileNameWithoutExtension(path).ToLowerInvariant();
        if (Regex.IsMatch(s,"股东协议|股東協議|shareholders?")) return "股东协议";
        if (Regex.IsMatch(s,"章程|articles.of.association|constitution")) return "章程";
        if (Regex.IsMatch(s,"股权转让|股份转让|share.purchase|equity.transfer")) return "股权转让";
        if (Regex.IsMatch(s,"投资协议|增资协议|增资认购|investment.agreement|subscription.agreement")) return "投资协议";
        return "";
    }
    public static string Clean(string path) {
        string s=Path.GetFileNameWithoutExtension(path).ToLowerInvariant();
        s=Regex.Replace(s,@"20\d{2}[-_.年]?\d{1,2}[-_.月]?\d{1,2}日?|\bv\d+(\.\d+)*|\d+","");
        s=Regex.Replace(s,"清洁版|修订版|定稿|初稿|草稿|终稿|最终版|签署版|修改版|最新版|新版|旧版|clean|final|draft|revised","");
        return Regex.Replace(s,@"[^\p{L}]+","");
    }
    public static double Score(string a,string b) {
        if (string.Equals(a,b,StringComparison.OrdinalIgnoreCase)) return -1;
        string ka=Kind(a),kb=Kind(b),x=Clean(a),y=Clean(b);
        if(ka!="" && kb!="" && ka!=kb) return -1;
        double sim=0;
        if(x.Length>0 && x==y) sim=1;
        else {
            var aa=new HashSet<string>();var bb=new HashSet<string>();
            for(int i=0;i<x.Length-1;i++) aa.Add(x.Substring(i,2));
            for(int i=0;i<y.Length-1;i++) bb.Add(y.Substring(i,2));
            if(aa.Count+bb.Count>0) sim=2.0*aa.Intersect(bb).Count()/(aa.Count+bb.Count);
        }
        return ka!="" && ka==kb ? .72+.28*sim : sim;
    }
    public static int[] Suggest(IList<string> old,IList<string> newer) {
        var answer=Enumerable.Repeat(-1,newer.Count).ToArray();
        if(old.Count==1 && newer.Count==1 && !string.Equals(old[0],newer[0],StringComparison.OrdinalIgnoreCase)) { answer[0]=0;return answer; }
        for(int n=0;n<newer.Count;n++) {
            var rank=Enumerable.Range(0,old.Count).Select(o=>new { Id=o,Value=Score(old[o],newer[n]) }).OrderByDescending(v=>v.Value).ToArray();
            if(rank.Length==0 || rank[0].Value<.58 || (rank.Length>1 && rank[0].Value-rank[1].Value<.10)) continue;
            int best=rank[0].Id;
            var reverse=Enumerable.Range(0,newer.Count).Select(j=>new {Id=j,Value=Score(old[best],newer[j])}).OrderByDescending(v=>v.Value).ToArray();
            if(reverse[0].Id==n && (reverse.Length==1 || reverse[0].Value-reverse[1].Value>=.10)) answer[n]=best;
        }
        return answer;
    }
}

