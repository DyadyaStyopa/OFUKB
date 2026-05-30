import SwiftUI

struct SidebarView: View {
    @EnvironmentObject private var appState: AppState

    var body: some View {
        List {
            Section {
                VStack(alignment: .leading, spacing: 8) {
                    Label(appState.statusText, systemImage: statusIcon)
                        .font(.headline)
                    if let code = appState.lastExitCode {
                        Text("Последний код: \(code)")
                            .foregroundStyle(.secondary)
                    } else {
                        Text("Выберите книгу, regnum и запустите обработку.")
                            .foregroundStyle(.secondary)
                    }
                }
                .padding(.vertical, 6)
            }

            Section("Действия") {
                Button {
                    appState.run()
                } label: {
                    Label("Запустить", systemImage: "play.fill")
                }
                .disabled(!appState.canRun)

                Button {
                    appState.stop()
                } label: {
                    Label("Остановить", systemImage: "stop.fill")
                }
                .disabled(!appState.isRunning)

                Button {
                    appState.openOutput()
                } label: {
                    Label("Открыть результат", systemImage: "doc.text.magnifyingglass")
                }

                Button {
                    appState.openOutputFolder()
                } label: {
                    Label("Открыть папку", systemImage: "folder")
                }
            }

            Section("Файл") {
                Text(appState.xlsxPath.isEmpty ? "Книга не выбрана" : URL(fileURLWithPath: appState.xlsxPath).lastPathComponent)
                    .foregroundStyle(appState.xlsxPath.isEmpty ? .secondary : .primary)
                    .lineLimit(2)
            }
        }
        .listStyle(.sidebar)
    }

    private var statusIcon: String {
        switch appState.status {
        case .idle:
            return "circle"
        case .running:
            return "clock"
        case .succeeded:
            return "checkmark.circle.fill"
        case .failed:
            return "exclamationmark.triangle.fill"
        case .stopped:
            return "stop.circle"
        }
    }
}
