import Foundation

extension Array where Element == String {
    var quotedCommand: String {
        map { item in
            if item.isEmpty {
                return "''"
            }
            let safe = CharacterSet(charactersIn: "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789_./:-")
            if item.unicodeScalars.allSatisfy({ safe.contains($0) }) {
                return item
            }
            return "'" + item.replacingOccurrences(of: "'", with: "'\\''") + "'"
        }
        .joined(separator: " ")
    }
}
