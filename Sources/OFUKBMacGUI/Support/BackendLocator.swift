import Foundation

enum BackendLocator {
    static let displayName = "OFUKB_CBR_PQ_alt_parser.py"
    static let sqliteExporterDisplayName = "cbr_sqlite_export.py"

    static func findBackend() -> URL? {
        findScript(displayName: displayName, resourceName: "OFUKB_CBR_PQ_alt_parser")
    }

    static func findSQLiteExporter() -> URL? {
        findScript(displayName: sqliteExporterDisplayName, resourceName: "cbr_sqlite_export")
    }

    private static func findScript(displayName: String, resourceName: String) -> URL? {
        if let bundled = Bundle.main.url(forResource: resourceName, withExtension: "py") {
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
