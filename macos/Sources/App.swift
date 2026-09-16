import AppKit
import SwiftUI
import UniformTypeIdentifiers

struct PairRow: Identifiable {
    let id = UUID()
    var new: URL
    var oldIndex: Int
    var selected = true
    var status: String
    var output: URL?
}

@MainActor final class RedlineModel: ObservableObject {
    @Published var oldFiles: [URL] = []
    @Published var newFiles: [URL] = []
    @Published var rows: [PairRow] = []
    @Published var running = false
    @Published var stopRequested = false
    @Published var message = "拖入文件只生成配对建议，确认后点击开始比对。"
    func add(_ urls: [URL], old: Bool) {
        guard !running else { return }
        var files = old ? oldFiles : newFiles
        var ignored = 0
        for url in urls {
            let clean = url.standardizedFileURL
            var isDirectory: ObjCBool = false
            guard ["docx", "doc", "docm"].contains(clean.pathExtension.lowercased()),
                  !clean.lastPathComponent.hasPrefix("~$"),
                  FileManager.default.fileExists(atPath: clean.path, isDirectory: &isDirectory), !isDirectory.boolValue else { ignored += 1; continue }
            if !files.contains(clean) { files.append(clean) }
        }
        if old { oldFiles = files } else { newFiles = files }
        refresh()
        if ignored > 0 { message += " 已忽略 \(ignored) 个不支持的文件。" }
    }
    func refresh() {
        let matches = Matcher.suggest(old: oldFiles, new: newFiles)
        rows = newFiles.indices.map { PairRow(new: newFiles[$0], oldIndex: matches[$0], status: matches[$0] < 0 ? "待手动配对" : "建议配对 · 请核对") }
        message = "旧版 \(oldFiles.count) 份 / 新版 \(newFiles.count) 份；建议 \(matches.filter { $0 >= 0 }.count) 对。确认后点击开始比对。"
    }
    func remove(_ url: URL, old: Bool) {
        guard !running else { return }
        if old { oldFiles.removeAll { $0 == url } } else { newFiles.removeAll { $0 == url } }
        refresh()
    }
    func choose(old: Bool) {
        let panel = NSOpenPanel()
        panel.allowsMultipleSelection = true
        panel.canChooseDirectories = false
        panel.allowedContentTypes = ["docx","doc","docm"].compactMap { UTType(filenameExtension: $0) }
        if panel.runModal() == .OK { add(panel.urls, old: old) }
    }
    func start() {
        guard !running else { return }
        let selected = rows.indices.filter { rows[$0].selected }
        guard !selected.isEmpty else { message = "请添加文件并勾选至少一组配对。"; return }
        var seen = Set<Int>()
        for i in selected {
            let o = rows[i].oldIndex
            guard oldFiles.indices.contains(o) else { message = "第 \(i+1) 行未配对，请选择旧版或取消勾选。"; return }
            guard seen.insert(o).inserted else { message = "同一旧版被多行使用，请调整为一对一配对。"; return }
            guard oldFiles[o] != rows[i].new else { message = "同一个文件不能与自己比较。"; return }
        }
        guard NSWorkspace.shared.urlForApplication(withBundleIdentifier: "com.microsoft.Word") != nil else {
            message = "未找到 Microsoft Word for Mac，请先安装并正常启动 Word 桌面版。"; return
        }
        guard let script = Bundle.main.url(forResource: "Compare", withExtension: "applescript") else { message = "缺少比对组件，请重新解压完整 App。"; return }
        running = true
        stopRequested = false
        for i in selected { rows[i].status = "等待比对"; rows[i].output = nil }
        Task {
            var succeeded = 0, failed = 0, completed = 0
            for i in selected {
                if stopRequested { rows[i].status = "未执行"; continue }
                let original = oldFiles[rows[i].oldIndex], revised = rows[i].new
                rows[i].status = "Word 正在比较…"
                message = "处理 \(completed+1) / \(selected.count)：\(revised.lastPathComponent)。若 Word 请求自动化或文件访问权限，请处理提示。"
                do {
                    let output = try await Task.detached { try WordWorker.run(old: original, new: revised, script: script) }.value
                    rows[i].output = output
                    rows[i].status = "完成"
                    succeeded += 1
                } catch {
                    failed += 1
                    rows[i].status = error.localizedDescription
                    if let failure = error as? WordWorker.Failure, failure.stopBatch { stopRequested = true }
                }
                completed += 1
            }
            running = false
            let skipped = selected.count-completed
            message = "成功 \(succeeded) 组，失败 \(failed) 组，未执行 \(skipped) 组。结果保存在各新版文件夹。"
        }
    }
}

struct FilePane: View {
    @ObservedObject var model: RedlineModel
    let old: Bool
    @State private var highlighted = false
    var files: [URL] { old ? model.oldFiles : model.newFiles }
    var body: some View {
        VStack(alignment: .leading, spacing: 8) {
            Text(old ? "旧版 · Original" : "新版 · Revised").font(.headline)
            ScrollView {
                VStack(alignment: .leading, spacing: 8) {
                    if files.isEmpty { Text("拖入一份或多份 Word 文件").foregroundColor(.secondary).padding(.vertical, 35) }
                    ForEach(files, id: \.self) { url in
                        HStack {
                            Text(url.lastPathComponent).lineLimit(1).truncationMode(.middle).help(url.path)
                            Spacer()
                            Button { model.remove(url, old: old) } label: { Image(systemName: "xmark.circle") }.buttonStyle(.borderless).disabled(model.running)
                        }
                    }
                }.frame(maxWidth: .infinity, alignment: .leading).padding(12)
            }
            .frame(height: 150)
            .background(highlighted ? Color.accentColor.opacity(0.12) : Color(NSColor.controlBackgroundColor))
            .cornerRadius(8)
            .overlay(RoundedRectangle(cornerRadius: 8).stroke(Color.secondary.opacity(0.3)))
            .onDrop(of: [UTType.fileURL.identifier], isTargeted: $highlighted) { providers in
                guard !model.running else { return false }
                for provider in providers {
                    _ = provider.loadObject(ofClass: URL.self) { url, _ in
                        if let url = url { Task { @MainActor in model.add([url], old: old) } }
                    }
                }
                return true
            }
            Button("添加文件…") { model.choose(old: old) }.disabled(model.running)
        }.frame(maxWidth: .infinity)
    }
}

struct MainView: View {
    @StateObject private var model = RedlineModel()
    var body: some View {
        VStack(alignment: .leading, spacing: 16) {
            Text("先确认配对，再开始比对").font(.system(size: 25, weight: .bold))
            Text("Word Redline 2.0 · macOS · Word 原生修订").foregroundColor(.secondary)
            HStack(alignment: .top, spacing: 20) { FilePane(model: model, old: true); FilePane(model: model, old: false) }
            Text("配对预览：可手动选择对应旧版；不需要的行取消勾选。").font(.callout)
            ScrollView {
                VStack(spacing: 0) {
                    HStack { Text("比对").frame(width: 38); Text("新版文件").frame(maxWidth: .infinity, alignment: .leading); Text("对应旧版").frame(width: 290, alignment: .leading); Text("状态 / 结果").frame(width: 180, alignment: .leading) }.font(.headline).padding(10)
                    Divider()
                    ForEach($model.rows) { $row in
                        HStack(spacing: 12) {
                            Toggle("", isOn: $row.selected).labelsHidden().frame(width: 28).disabled(model.running)
                            Text(row.new.lastPathComponent).lineLimit(2).frame(maxWidth: .infinity, alignment: .leading).help(row.new.path)
                            Picker("对应旧版", selection: $row.oldIndex) {
                                Text("— 请选择旧版 —").tag(-1)
                                ForEach(model.oldFiles.indices, id: \.self) { i in Text("\(i+1). \(model.oldFiles[i].lastPathComponent)").tag(i) }
                            }.labelsHidden().frame(width: 290).disabled(model.running)
                            VStack(alignment: .leading, spacing: 4) {
                                Text(row.status).font(.caption).lineLimit(3).help(row.status)
                                if let output = row.output { Button("打开结果") { NSWorkspace.shared.open(output) }.help(output.path) }
                            }.frame(width: 180, alignment: .leading)
                        }.padding(10).background(row.oldIndex < 0 ? Color.orange.opacity(0.08) : Color.clear)
                        Divider()
                    }
                }
            }.frame(minHeight: 180).background(Color(NSColor.controlBackgroundColor)).cornerRadius(8)
            Text(model.message).font(.callout).foregroundColor(.secondary).textSelection(.enabled).frame(minHeight: 38, alignment: .leading)
            HStack {
                Button("开始比对") { model.start() }.buttonStyle(.borderedProminent).disabled(model.running)
                Button("停止后续比对") { model.stopRequested = true; model.message = "当前组完成后停止；已生成的结果保留。" }.disabled(!model.running || model.stopRequested)
                Button("清空 / 下一组") { model.oldFiles = []; model.newFiles = []; model.refresh() }.disabled(model.running)
                Spacer()
                if model.running { ProgressView().controlSize(.small) }
            }
        }.padding(24).frame(minWidth: 980, minHeight: 700)
    }
}

struct RedlineApp: App {
    var body: some Scene { WindowGroup("Word Redline 2.0") { MainView() }.commands { CommandGroup(replacing: .newItem) {} } }
}

@main enum EntryPoint {
    @MainActor
    static func main() {
        if CommandLine.arguments.contains("--self-test") { Matcher.selfTest(); return }
        RedlineApp.main()
    }
}
