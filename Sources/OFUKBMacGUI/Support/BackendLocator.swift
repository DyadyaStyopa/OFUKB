import Foundation

enum BackendLocator {
    static let displayName = "OFUKB_CBR_PQ_alt_parser.py"

    static func findBackend() -> URL? {
        if let bundled = Bundle.main.url(
            forResource: "OFUKB_CBR_PQ_alt_parser",
            withExtension: "py"
        ) {
            return bundled
        }

        let cwd = URL(fileURLWithPath: FileManager.default.currentDirectoryPath)
        let local = cwd.appendingPathComponent(displayName)
        if FileManager.default.fileExists(atPath: local.path) {
            return local
        }

        let executable = Bundle.main.bundleURL
            .deletingLastPathComponent()
            .deletingLastPathComponent()
            .appendingPathComponent(displayName)
        if FileManager.default.fileExists(atPath: executable.path) {
            return executable
        }

        return nil
    }
}
