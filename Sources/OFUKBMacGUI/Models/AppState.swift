import AppKit
import Foundation

@MainActor
final class AppState: ObservableObject {
    @Published var xlsxPath = ""
    @Published var outputPath = ""
    @Published var regnum = ""
    @Published var cacheDir = ""
    @Published var sqliteSource = ""
    @Published var logFile = ""
    @Published var debugDir = ""

    @Published var verbose = true
    @Published var debug = false
    @Published var noCache = false
    @Published var dumpM = false
    @Published var listOnly = false
    @Published var sqliteAllBanks = false
    @Published var fillFromSQLite = false

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

    var canRunTransform: Bool {
        canRun && sqliteAllBanks
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
        if let command = activeCommand() {
            return command.displayCommand
        }
        return (["python3", activeScriptDisplayName] + backendArguments()).quotedCommand
    }

    private var activeScriptDisplayName: String {
        sqliteAllBanks ? BackendLocator.sqliteExporterDisplayName : BackendLocator.displayName
    }

    private var activeModeName: String {
        sqliteAllBanks ? "sqlite" : "excel"
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

    func chooseSQLiteSource() {
        let panel = NSOpenPanel()
        panel.title = "Выберите SQLite-базу"
        panel.allowsMultipleSelection = false
        panel.canChooseDirectories = false
        panel.canChooseFiles = true
        panel.begin { [weak self] response in
            guard response == .OK, let url = panel.url else { return }
            Task { @MainActor in
                self?.sqliteSource = url.path
            }
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
        if sqliteAllBanks {
            fillFromSQLite = false
        }
        refreshOutputPath()
    }

    func toggleFillFromSQLite() {
        if fillFromSQLite {
            sqliteAllBanks = false
        }
        if outputUsesAutoPath || outputPath.isEmpty {
            outputPath = defaultOutputPath()
            outputUsesAutoPath = true
        }
    }

    func run(transformOnly: Bool = false) {
        guard !isRunning else { return }
        if transformOnly && !sqliteAllBanks {
            appendLog("Transform из HTML-кэша доступен только для SQLite по всем действующим банкам.\n")
            status = .failed(1)
            return
        }
        guard let command = activeCommand(transformOnly: transformOnly) else {
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
        if fillFromSQLite {
            let trimmedSQLite = sqliteSource.trimmingCharacters(in: .whitespacesAndNewlines)
            guard !trimmedSQLite.isEmpty else {
                appendLog("Выберите SQLite-базу для заполнения XLSX.\n")
                status = .failed(1)
                return
            }
            guard FileManager.default.fileExists(atPath: trimmedSQLite) else {
                appendLog("SQLite-база не найдена: \(trimmedSQLite)\n")
                status = .failed(1)
                return
            }
            guard !regnum.trimmingCharacters(in: .whitespacesAndNewlines).isEmpty else {
                appendLog("Укажите regnum для выборки из SQLite.\n")
                status = .failed(1)
                return
            }
        }

        clearLog()
        appendLog("Команда:\n\(command.displayCommand)\n\n")
        status = .running
        isRunning = true
        lastExitCode = nil

        let workDir = URL(fileURLWithPath: trimmedXLSX).deletingLastPathComponent()
        let newRunner = PythonRunner()
        runner = newRunner
        newRunner.run(
            executableURL: command.executableURL,
            arguments: command.arguments,
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

    private func backendArguments(transformOnly: Bool = false) -> [String] {
        if sqliteAllBanks {
            var args: [String] = [xlsxPath, "--all-banks", transformOnly ? "--transform-only" : "--prefetch", "--workers", "4"]
            if !outputPath.isEmpty {
                args += ["--output", outputPath]
            }
            args.append("--replace")
            if verbose { args.append("--verbose") }
            if !cacheDir.isEmpty { args += ["--cache-dir", cacheDir] }
            return args
        }

        var args: [String] = [xlsxPath]
        if fillFromSQLite {
            if !sqliteSource.isEmpty {
                args += ["--from-sqlite", sqliteSource]
            }
            if !regnum.isEmpty {
                args += ["--sqlite-regnum", regnum]
            }
            if !outputPath.isEmpty {
                args += ["--output", outputPath]
            }
            if verbose { args.append("--verbose") }
            return args
        }
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
        if fillFromSQLite {
            let suffix = regnum.isEmpty ? "_sqlite_filled.xlsx" : "_regnum_\(regnum)_sqlite_filled.xlsx"
            return url.deletingPathExtension().path + suffix
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

    private func activeCommand(transformOnly: Bool = false) -> RunnerCommand? {
        if let cli = BackendLocator.findBundledCLI() {
            let args = [activeModeName] + backendArguments(transformOnly: transformOnly)
            return RunnerCommand(
                executableURL: cli,
                arguments: args,
                displayCommand: ([cli.path] + args).quotedCommand
            )
        }

        guard let script = activeScriptURL() else { return nil }
        let args = [script.path] + backendArguments(transformOnly: transformOnly)
        return RunnerCommand(
            executableURL: URL(fileURLWithPath: "/usr/bin/env"),
            arguments: ["python3"] + args,
            displayCommand: (["python3"] + args).quotedCommand
        )
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

private struct RunnerCommand {
    let executableURL: URL
    let arguments: [String]
    let displayCommand: String
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
