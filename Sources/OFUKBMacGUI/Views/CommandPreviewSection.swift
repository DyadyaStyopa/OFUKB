import SwiftUI

struct CommandPreviewSection: View {
    @EnvironmentObject private var appState: AppState

    var body: some View {
        GroupBox {
            HStack(spacing: 12) {
                Text(appState.commandPreview)
                    .font(.system(.body, design: .monospaced))
                    .lineLimit(2)
                    .textSelection(.enabled)
                    .frame(maxWidth: .infinity, alignment: .leading)

                Button {
                    NSPasteboard.general.clearContents()
                    NSPasteboard.general.setString(appState.commandPreview, forType: .string)
                } label: {
                    Label("Копировать", systemImage: "doc.on.doc")
                }
            }
            .padding(6)
        } label: {
            Label("Команда backend", systemImage: "terminal")
        }
    }
}
