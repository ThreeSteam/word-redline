import Foundation

enum Matcher {
    static func replace(_ text: String, _ pattern: String) -> String {
        text.replacingOccurrences(of: pattern, with: "", options: .regularExpression)
    }
    static func kind(_ url: URL) -> String {
        let name = url.deletingPathExtension().lastPathComponent.lowercased()
        for (pattern, label) in [
            ("股东协议|股東協議|shareholders?", "股东协议"),
            ("章程|articles.of.association|constitution", "章程"),
            ("股权转让|股份转让|share.purchase|equity.transfer", "股权转让"),
            ("投资协议|增资协议|增资认购|investment.agreement|subscription.agreement", "投资协议")
        ] where name.range(of: pattern, options: .regularExpression) != nil { return label }
        return ""
    }
    static func clean(_ url: URL) -> String {
        var name = url.deletingPathExtension().lastPathComponent.lowercased()
        name = replace(name, #"20\d{2}[-_.年]?\d{1,2}[-_.月]?\d{1,2}日?|\bv\d+(\.\d+)*|\d+"#)
        name = replace(name, "清洁版|修订版|定稿|初稿|草稿|终稿|最终版|签署版|修改版|最新版|新版|旧版|clean|final|draft|revised")
        return replace(name, #"[^\p{L}]+"#)
    }
    static func score(_ old: URL, _ new: URL) -> Double {
        if old.standardizedFileURL == new.standardizedFileURL { return -1 }
        let a = kind(old), b = kind(new), x = clean(old), y = clean(new)
        if !a.isEmpty && !b.isEmpty && a != b { return -1 }
        func bigrams(_ text: String) -> Set<String> {
            let c = Array(text)
            guard c.count > 1 else { return [] }
            return Set((0..<(c.count-1)).map { String(c[$0...($0+1)]) })
        }
        let xx = bigrams(x), yy = bigrams(y)
        let similarity: Double = !x.isEmpty && x == y ? 1 :
            (xx.count+yy.count == 0 ? 0 : 2 * Double(xx.intersection(yy).count) / Double(xx.count+yy.count))
        return !a.isEmpty && a == b ? 0.72 + 0.28 * similarity : similarity
    }
    static func suggest(old: [URL], new: [URL]) -> [Int] {
        if old.count == 1 && new.count == 1 && old[0].standardizedFileURL != new[0].standardizedFileURL { return [0] }
        return new.indices.map { n in
            let ranked = old.indices.map { ($0, score(old[$0], new[n])) }.sorted { $0.1 > $1.1 }
            guard let best = ranked.first, best.1 >= 0.58 else { return -1 }
            if ranked.count > 1 && best.1 - ranked[1].1 < 0.10 { return -1 }
            let reverse = new.indices.map { ($0, score(old[best.0], new[$0])) }.sorted { $0.1 > $1.1 }
            guard reverse.first?.0 == n else { return -1 }
            if reverse.count > 1 && reverse[0].1 - reverse[1].1 < 0.10 { return -1 }
            return best.0
        }
    }
    static func selfTest() {
        func urls(_ names: [String]) -> [URL] { names.map { URL(fileURLWithPath: "/test/" + $0) } }
        let old = urls(["01甲公司投资协议A轮定稿20260706.docx", "02甲公司股东协议.docx", "03甲公司章程.docx"])
        let new = urls(["新公司章程修订版0914.docx", "新公司增资协议2026.docx", "新公司股东协议最终版.docx"])
        precondition(suggest(old: old, new: new) == [2,0,1], "Shuffled keyword pairs")
        precondition(suggest(old: urls(["甲投资协议.docx", "乙投资协议.docx"]), new: urls(["丙投资协议.docx"])) == [-1], "Ambiguous pairing")
        precondition(suggest(old: urls(["old.docx"]), new: urls(["new.docx"])) == [0], "Single pair")
        precondition(suggest(old: urls(["same.docx"]), new: urls(["same.docx"])) == [-1], "Same file")
        precondition(score(old[0], old[1]) < 0, "Conflicting kinds")
        precondition(suggest(old: urls(["甲章程.docx"]), new: urls(["甲章程修订版.docx", "甲章程最终版.docx"])) == [-1,-1], "Duplicate candidates")
        print("PASS: six filename matching tests")
    }
}
