import AppKit
import SwiftUI

struct HeaderView: View {
    @EnvironmentObject private var appState: AppState

    var body: some View {
        HStack(spacing: 14) {
            Image(nsImage: headerIcon)
                .resizable()
                .frame(width: 42, height: 42)
                .clipShape(RoundedRectangle(cornerRadius: 8))

            VStack(alignment: .leading, spacing: 2) {
                Text("ОФУКБ ЦБ Power Query")
                    .font(.title2.weight(.semibold))
                Text("Нативная macOS-оболочка для заполнения Excel-книги через Python backend")
                    .foregroundStyle(.secondary)
            }

            Spacer()

            if appState.isRunning {
                ProgressView()
                    .controlSize(.small)
            }

            Button {
                appState.run()
            } label: {
                Label("Запустить", systemImage: "play.fill")
            }
            .keyboardShortcut(.return, modifiers: [.command])
            .buttonStyle(.borderedProminent)
            .disabled(!appState.canRun)

            Button {
                appState.stop()
            } label: {
                Label("Стоп", systemImage: "stop.fill")
            }
            .disabled(!appState.isRunning)
        }
        .padding(.horizontal, 24)
        .padding(.vertical, 16)
    }

    private var headerIcon: NSImage {
        if let url = Bundle.main.url(forResource: "app_icon", withExtension: "png"),
           let image = NSImage(contentsOf: url) {
            return image
        }
        return NSImage(named: NSImage.applicationIconName) ?? NSImage()
    }
}
