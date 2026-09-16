import Foundation

enum WordWorker {
    struct Failure: LocalizedError {
        let message: String
        let stopBatch: Bool
        var errorDescription: String? { message }
    }
    static func run(old: URL, new: URL, script: URL) throws -> URL {
        let fm = FileManager.default
        let workspace = fm.temporaryDirectory.appendingPathComponent("WordRedline-" + UUID().uuidString, isDirectory: true)
        try fm.createDirectory(at: workspace, withIntermediateDirectories: true)
        // Only disposable copies are opened in the user's existing Word session.
        let oldCopy = workspace.appendingPathComponent("Original." + old.pathExtension)
        let newCopy = workspace.appendingPathComponent("Revised." + new.pathExtension)
        let generated = workspace.appendingPathComponent("Comparison.docx")
        try fm.copyItem(at: old, to: oldCopy)
        try fm.copyItem(at: new, to: newCopy)
        let process = Process(), output = Pipe(), errors = Pipe()
        process.executableURL = URL(fileURLWithPath: "/usr/bin/osascript")
        process.arguments = [script.path, oldCopy.path, newCopy.path, generated.path]
        process.standardOutput = output
        process.standardError = errors
        let complete = DispatchSemaphore(value: 0)
        process.terminationHandler = { _ in complete.signal() }
        do { try process.run() } catch {
            try? fm.removeItem(at: workspace)
            throw error
        }
        if complete.wait(timeout: .now() + 660) == .timedOut {
            process.terminate()
            throw Failure(message: "Word 响应超时，已停止后续任务。请检查 Word 中的授权或提示窗口；临时副本保留在 \(workspace.path)。", stopBatch: true)
        }
        let detail = String(data: errors.fileHandleForReading.readDataToEndOfFile(), encoding: .utf8) ?? ""
        _ = output.fileHandleForReading.readDataToEndOfFile()
        guard process.terminationStatus == 0 && fm.fileExists(atPath: generated.path) else {
            throw Failure(message: "Word 比对失败：\(detail.trimmingCharacters(in: .whitespacesAndNewlines))。临时副本：\(workspace.path)", stopBatch: true)
        }
        // Validate the saved file is an Open XML document rather than a renamed .doc.
        let check = Process(), listing = Pipe()
        check.executableURL = URL(fileURLWithPath: "/usr/bin/unzip")
        check.arguments = ["-Z1", generated.path]
        check.standardOutput = listing
        check.standardError = Pipe()
        try check.run()
        let entries = String(data: listing.fileHandleForReading.readDataToEndOfFile(), encoding: .utf8) ?? ""
        check.waitUntilExit()
        guard check.terminationStatus == 0 && entries.components(separatedBy: .newlines).contains("word/document.xml") else {
            throw Failure(message: "Word 未保存为有效的 DOCX。结果暂存在 \(workspace.path)，未写入新版目录。", stopBatch: true)
        }
        let directory = new.deletingLastPathComponent(), stem = new.deletingPathExtension().lastPathComponent
        for index in 1...10000 {
            let suffix = index == 1 ? " - redline" : " - redline (\(index))"
            let destination = directory.appendingPathComponent(stem + suffix + ".docx")
            do {
                // copyItem fails if a destination already exists, including concurrent writes.
                try fm.copyItem(at: generated, to: destination)
                try? fm.removeItem(at: workspace)
                return destination
            } catch let error as NSError {
                if error.domain == NSCocoaErrorDomain && error.code == NSFileWriteFileExistsError { continue }
                throw Failure(message: "保存失败：\(error.localizedDescription)。比较结果保留在 \(generated.path)", stopBatch: false)
            }
        }
        throw Failure(message: "同名结果过多。比较结果保留在 \(generated.path)", stopBatch: false)
    }
}
