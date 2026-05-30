import SwiftUI

struct SettingsView: View {
    @EnvironmentObject private var appState: AppState

    var body: some View {
        Form {
            Toggle("Подробный лог по умолчанию", isOn: $appState.verbose)
            Toggle("Debug-режим по умолчанию", isOn: $appState.debug)
        }
        .padding(24)
        .frame(width: 380)
    }
}
