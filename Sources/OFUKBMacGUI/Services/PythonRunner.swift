import Foundation

@MainActor
final class PythonRunner {
    private var process: Process?
    private var pipe: Pipe?

    func run(
        executableURL: URL,
        arguments: [String],
        workingDirectory: URL,
        output: @escaping (String) -> Void,
        completion: @escaping (Int32) -> Void
    ) {
        let process = Process()
        let pipe = Pipe()

        process.executableURL = executableURL
        process.arguments = arguments
        process.currentDirectoryURL = workingDirectory
        process.standardOutput = pipe
        process.standardError = pipe

        pipe.fileHandleForReading.readabilityHandler = { handle in
            let data = handle.availableData
            guard !data.isEmpty, let text = String(data: data, encoding: .utf8) else { return }
            Task { @MainActor in
                output(text)
            }
        }

        process.terminationHandler = { [weak self] proc in
            Task { @MainActor in
                self?.pipe?.fileHandleForReading.readabilityHandler = nil
                self?.process = nil
                self?.pipe = nil
                completion(proc.terminationStatus)
            }
        }

        self.process = process
        self.pipe = pipe

        do {
            try process.run()
        } catch {
            pipe.fileHandleForReading.readabilityHandler = nil
            self.process = nil
            self.pipe = nil
            output("Не удалось запустить процесс: \(error.localizedDescription)\n")
            completion(1)
        }
    }

    func stop() {
        process?.terminate()
        process = nil
        pipe?.fileHandleForReading.readabilityHandler = nil
        pipe = nil
    }
}
