import SwiftUI

struct AdvancedPathsSection: View {
    @EnvironmentObject private var appState: AppState
    @State private var isExpanded = false

    var body: some View {
        DisclosureGroup(isExpanded: $isExpanded) {
            VStack(alignment: .leading, spacing: 12) {
                PathField(
                    title: "HTML-кэш",
                    value: appState.cacheDir,
                    placeholder: "По умолчанию backend использует pq_html_cache",
                    systemImage: "externaldrive",
                    buttonTitle: "Выбрать",
                    action: appState.chooseCacheDir
                )
                PathField(
                    title: "Файл лога",
                    value: appState.logFile,
                    placeholder: "По умолчанию рядом с книгой при verbose/debug",
                    systemImage: "doc.plaintext",
                    buttonTitle: "Выбрать",
                    action: appState.chooseLogFile
                )
                PathField(
                    title: "Debug-папка",
                    value: appState.debugDir,
                    placeholder: "По умолчанию рядом с книгой при debug",
                    systemImage: "folder.badge.gearshape",
                    buttonTitle: "Выбрать",
                    action: appState.chooseDebugDir
                )
            }
            .padding(.top, 12)
        } label: {
            Label("Дополнительные пути", systemImage: "slider.horizontal.3")
        }
    }
}
