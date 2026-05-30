import AppKit
import SwiftUI

@main
struct OFUKBMacGUIApp: App {
    @NSApplicationDelegateAdaptor(AppDelegate.self) private var appDelegate
    @StateObject private var appState = AppState()

    var body: some Scene {
        WindowGroup("ОФУКБ ЦБ Power Query") {
            ContentView()
                .environmentObject(appState)
                .frame(minWidth: 980, minHeight: 680)
        }
        .commands {
            CommandGroup(replacing: .newItem) {}
            CommandMenu("Запуск") {
                Button(appState.isRunning ? "Остановить" : "Запустить") {
                    if appState.isRunning {
                        appState.stop()
                    } else {
                        appState.run()
                    }
                }
                .keyboardShortcut(.return, modifiers: [.command])

                Button("Очистить лог") {
                    appState.clearLog()
                }
                .keyboardShortcut("k", modifiers: [.command])
            }
        }

        Settings {
            SettingsView()
                .environmentObject(appState)
        }
    }
}

final class AppDelegate: NSObject, NSApplicationDelegate {
    func applicationDidFinishLaunching(_ notification: Notification) {
        NSApp.setActivationPolicy(.regular)
        NSApp.activate(ignoringOtherApps: true)
    }
}
