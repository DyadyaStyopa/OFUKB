import SwiftUI

struct ContentView: View {
    @EnvironmentObject private var appState: AppState

    var body: some View {
        NavigationSplitView {
            SidebarView()
                .navigationSplitViewColumnWidth(min: 230, ideal: 260, max: 310)
        } detail: {
            VStack(spacing: 0) {
                HeaderView()
                Divider()
                ScrollView {
                    VStack(alignment: .leading, spacing: 18) {
                        WorkbookSection()
                        RunOptionsSection()
                        AdvancedPathsSection()
                        CommandPreviewSection()
                    }
                    .padding(24)
                    .frame(maxWidth: .infinity, alignment: .leading)
                }
                Divider()
                LogSection()
            }
        }
    }
}
