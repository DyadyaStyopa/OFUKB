import AppKit
import Foundation

@MainActor
final class AppState: ObservableObject {
    @Published var xlsxPath = ""
    @Published var outputPath = ""
    @Published var regnum = ""
    @Published var cacheDir = ""
    @Published var logFile = ""
    @Published var debugDir = ""

    @Published var verbose = true
    @Published var debug = false
    @Published var noCache = false
    @Published var dumpM = false
    @Published var listOnly = false
    @Published var sqliteAllBanks = false

    @Published var logText = ""
    @Published var status = RunStatus.idle
    @Published var lastExitCode: Int32?
    @Published var isRunning = false
    @Published var selectedLogLevel = LogLevel.normal

    private var runner: PythonRunner?
    private var outputUsesAutoPath = true

    var canRun: Bool {
        !xlsxPath.trimmingCharacters(in: .whitespacesAndNewlines).isEmpty && !isRunning
    }

    var statusText: String {
        switch status {
        case .idle:
            return "Готово к запуску"
        case .running:
            return "Выполняется"
        case .succeeded:
            return "Готово"
        case .failed(let code):
            return "Ошибка, код \(code)"
        case .stopped:
            return "Остановлено"
        }
    }

    var commandPreview: String {
        (["python3", activeScriptDisplayName] + backendArguments()).quotedCommand
    }

    private var activeScriptDisplayName: String {
        sqliteAllBanks ? BackendLocator.sqliteExporterDisplayName : BackendLocator.displayName
    }

    func chooseXLSX() {
        let panel = NSOpenPanel()
        panel.title = "Выберите Excel-книгу"
        panel.allowedContentTypes = [.xlsx]
        panel.allowsMultipleSelection = false
        panel.canChooseDirectories = false
        panel.begin { [weak self] response in
            guard response == .OK, let url = panel.url else { return }
            Task { @MainActor in
                self?.setXLSXPath(url.path)
            }
        }
    }

    func chooseOutput() {
        let panel = NSSavePanel()
        panel.title = "Куда сохранить результат"
        if !sqliteAllBanks {
            panel.allowedContentTypes = [.xlsx]
        }
        panel.nameFieldStringValue = defaultOutputName()
        panel.begin { [weak self] response in
            guard response == .OK, let url = panel.url else { return }
            Task { @MainActor in
                self?.outputPath = url.path
                self?.outputUsesAutoPath = false
            }
        }
    }

    func chooseCacheDir() {
        chooseDirectory(title: "Выберите папку HTML-кэша") { [weak self] path in
            self?.cacheDir = path
        }
    }

    func chooseDebugDir() {
        chooseDirectory(title: "Выберите debug-папку") { [weak self] path in
            self?.debugDir = path
        }
    }

    func chooseLogFile() {
        let panel = NSSavePanel()
        panel.title = "Куда сохранить лог"
        panel.allowedContentTypes = [.plainText]
        panel.nameFieldStringValue = "ofukb_run_log.txt"
        panel.begin { [weak self] response in
            guard response == .OK, let url = panel.url else { return }
            Task { @MainActor in
                self?.logFile = url.path
            }
        }
    }

    func setXLSXPath(_ path: String) {
        xlsxPath = path
        if outputUsesAutoPath || outputPath.isEmpty {
            outputPath = defaultOutputPath()
            outputUsesAutoPath = true
        }
    }

    func refreshOutputPath() {
        if outputUsesAutoPath || outputPath.isEmpty {
            outputPath = defaultOutputPath()
            outputUsesAutoPath = true
        }
    }

    func toggleSQLiteAllBanks() {
        if outputUsesAutoPath || outputPath.isEmpty {
            outputPath = defaultOutputPath()
            outputUsesAutoPath = true
        }
    }

    func run() {
        guard !isRunning else { return }
        guard let backend = activeScriptURL() else {
            appendLog("Не найден \(activeScriptDisplayName). Положите скрипт рядом с приложением или в корень репозитория.\n")
            status = .failed(1)
            return
        }
        let trimmedXLSX = xlsxPath.trimmingCharacters(in: .whitespacesAndNewlines)
        guard !trimmedXLSX.isEmpty else {
            appendLog("Выберите исходный .xlsx файл.\n")
            status = .failed(1)
            return
        }
        guard FileManager.default.fileExists(atPath: trimmedXLSX) else {
            appendLog("Excel-файл не найден: \(trimmedXLSX)\n")
            status = .failed(1)
            return
        }

        clearLog()
        let args = [backend.path] + backendArguments()
        appendLog("Команда:\n\(("python3 " + args.quotedCommand))\n\n")
        status = .running
        isRunning = true
        lastExitCode = nil

        let workDir = URL(fileURLWithPath: trimmedXLSX).deletingLastPathComponent()
        let newRunner = PythonRunner()
        runner = newRunner
        newRunner.run(
            arguments: args,
            workingDirectory: workDir,
            output: { [weak self] text in
                self?.appendLog(text)
            },
            completion: { [weak self] code in
                self?.runner = nil
                self?.lastExitCode = code
                self?.isRunning = false
                self?.status = code == 0 ? .succeeded : .failed(code)
                self?.appendLog("\n[Процесс завершен с кодом \(code)]\n")
            }
        )
    }

    func stop() {
        guard isRunning else { return }
        runner?.stop()
        runner = nil
        isRunning = false
        status = .stopped
        appendLog("\n[Процесс остановлен пользователем]\n")
    }

    func clearLog() {
        logText = ""
    }

    func appendLog(_ text: String) {
        logText += text
    }

    func openOutput() {
        let path = outputPath.isEmpty ? defaultOutputPath() : outputPath
        openURLIfExists(path: path, fallbackMessage: "Файл результата пока не найден")
    }

    func openOutputFolder() {
        let path = outputPath.isEmpty ? xlsxPath : outputPath
        guard !path.isEmpty else { return }
        let url = URL(fileURLWithPath: path)
        let folder = url.hasDirectoryPath ? url : url.deletingLastPathComponent()
        NSWorkspace.shared.open(folder)
    }

    private func backendArguments() -> [String] {
        if sqliteAllBanks {
            var args: [String] = [xlsxPath, "--all-banks"]
            if !outputPath.isEmpty {
                args += ["--output", outputPath]
            }
            args.append("--replace")
            if noCache { args.append("--no-cache") }
            if verbose { args.append("--verbose") }
            if !cacheDir.isEmpty { args += ["--cache-dir", cacheDir] }
            return args
        }

        var args: [String] = [xlsxPath]
        if !outputPath.isEmpty && !listOnly {
            args += ["--output", outputPath]
        }
        if !regnum.isEmpty {
            args += ["--regnum", regnum]
        }
        if listOnly { args.append("--list") }
        if dumpM { args.append("--dump-m") }
        if noCache { args.append("--no-cache") }
        if verbose { args.append("--verbose") }
        if debug { args.append("--debug") }
        if !cacheDir.isEmpty { args += ["--cache-dir", cacheDir] }
        if !logFile.isEmpty { args += ["--log-file", logFile] }
        if !debugDir.isEmpty { args += ["--debug-dir", debugDir] }
        return args
    }

    private func defaultOutputPath() -> String {
        guard !xlsxPath.isEmpty else { return "" }
        let url = URL(fileURLWithPath: xlsxPath)
        if sqliteAllBanks {
            return url.deletingPathExtension().path + "_all_active_banks.sqlite"
        }
        let suffix = regnum.isEmpty ? "_python_filled.xlsx" : "_regnum_\(regnum)_python_filled.xlsx"
        return url.deletingPathExtension().path + suffix
    }

    private func defaultOutputName() -> String {
        guard !xlsxPath.isEmpty else { return sqliteAllBanks ? "cbr_banks.sqlite" : "result.xlsx" }
        return URL(fileURLWithPath: defaultOutputPath()).lastPathComponent
    }

    private func activeScriptURL() -> URL? {
        sqliteAllBanks ? BackendLocator.findSQLiteExporter() : BackendLocator.findBackend()
    }

    private func chooseDirectory(title: String, completion: @escaping (String) -> Void) {
        let panel = NSOpenPanel()
        panel.title = title
        panel.canChooseFiles = false
        panel.canChooseDirectories = true
        panel.allowsMultipleSelection = false
        panel.begin { response in
            guard response == .OK, let url = panel.url else { return }
            Task { @MainActor in
                completion(url.path)
            }
        }
    }

    private func openURLIfExists(path: String, fallbackMessage: String) {
        guard !path.isEmpty else { return }
        let url = URL(fileURLWithPath: path)
        if FileManager.default.fileExists(atPath: url.path) {
            NSWorkspace.shared.open(url)
        } else {
            appendLog("\(fallbackMessage): \(path)\n")
        }
    }
}

enum RunStatus: Equatable {
    case idle
    case running
    case succeeded
    case failed(Int32)
    case stopped
}

enum LogLevel: String, CaseIterable, Identifiable {
    case normal = "Обычный"
    case debug = "Debug"

    var id: String { rawValue }
}
