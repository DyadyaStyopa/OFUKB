import SwiftUI

struct RunOptionsSection: View {
    @EnvironmentObject private var appState: AppState

    var body: some View {
        GroupBox {
            Grid(alignment: .leading, horizontalSpacing: 28, verticalSpacing: 12) {
                GridRow {
                    Toggle("Подробный лог", isOn: $appState.verbose)
                    Toggle("Debug-режим", isOn: $appState.debug)
                    Toggle("Не использовать кэш", isOn: $appState.noCache)
                }
                GridRow {
                    Toggle("Сохранить M-код", isOn: $appState.dumpM)
                    Toggle("Только список таблиц", isOn: $appState.listOnly)
                    EmptyView()
                }
            }
            .toggleStyle(.checkbox)
            .padding(6)
        } label: {
            Label("Режимы запуска", systemImage: "switch.2")
        }
    }
}
