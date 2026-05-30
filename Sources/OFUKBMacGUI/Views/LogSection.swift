import SwiftUI

struct LogSection: View {
    @EnvironmentObject private var appState: AppState

    var body: some View {
        VStack(spacing: 0) {
            HStack {
                Label("Лог выполнения", systemImage: "text.alignleft")
                    .font(.headline)
                Spacer()
                Button {
                    appState.clearLog()
                } label: {
                    Label("Очистить", systemImage: "trash")
                }
            }
            .padding(.horizontal, 18)
            .padding(.vertical, 10)

            ScrollViewReader { proxy in
                ScrollView {
                    Text(appState.logText.isEmpty ? "Лог появится после запуска." : appState.logText)
                        .font(.system(.body, design: .monospaced))
                        .textSelection(.enabled)
                        .frame(maxWidth: .infinity, alignment: .leading)
                        .padding(14)
                        .id("log-end")
                }
                .background(Color.secondary.opacity(0.07))
                .onChangeCompat(of: appState.logText) {
                    proxy.scrollTo("log-end", anchor: .bottom)
                }
            }
        }
        .frame(minHeight: 220, idealHeight: 260)
    }
}
