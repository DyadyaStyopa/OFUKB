import SwiftUI
import UniformTypeIdentifiers

struct WorkbookSection: View {
    @EnvironmentObject private var appState: AppState

    var body: some View {
        GroupBox {
            VStack(alignment: .leading, spacing: 14) {
                FileDropZone(isTargeted: false)
                    .onDrop(of: [.fileURL], isTargeted: nil) { providers in
                        handleDrop(providers)
                    }

                PathField(
                    title: "Excel-книга",
                    value: appState.xlsxPath,
                    placeholder: "Выберите исходный .xlsx файл",
                    systemImage: "tablecells",
                    buttonTitle: "Выбрать",
                    action: appState.chooseXLSX
                )

                HStack(alignment: .firstTextBaseline, spacing: 12) {
                    Text("regnum")
                        .frame(width: 110, alignment: .leading)
                        .foregroundStyle(.secondary)
                    TextField("Например 1000", text: $appState.regnum)
                        .textFieldStyle(.roundedBorder)
                        .frame(width: 150)
                        .onChangeCompat(of: appState.regnum) {
                            appState.refreshOutputPath()
                        }
                    Text("Пусто = использовать номер из M-кода")
                        .foregroundStyle(.secondary)
                }

                PathField(
                    title: "Результат",
                    value: appState.outputPath,
                    placeholder: "Путь сформируется автоматически",
                    systemImage: "square.and.arrow.down",
                    buttonTitle: "Куда сохранить",
                    action: appState.chooseOutput
                )
            }
            .padding(6)
        } label: {
            Label("Исходные данные", systemImage: "doc.badge.gearshape")
        }
    }

    private func handleDrop(_ providers: [NSItemProvider]) -> Bool {
        guard let provider = providers.first else { return false }
        provider.loadItem(forTypeIdentifier: UTType.fileURL.identifier, options: nil) { item, _ in
            guard
                let data = item as? Data,
                let url = URL(dataRepresentation: data, relativeTo: nil),
                url.pathExtension.lowercased() == "xlsx"
            else { return }
            Task { @MainActor in
                appState.setXLSXPath(url.path)
            }
        }
        return true
    }
}

struct FileDropZone: View {
    let isTargeted: Bool

    var body: some View {
        VStack(spacing: 8) {
            Image(systemName: "arrow.down.doc")
                .font(.system(size: 28))
                .foregroundStyle(.secondary)
            Text("Перетащите .xlsx сюда или выберите файл")
                .font(.headline)
            Text("Исходная книга не изменяется: результат записывается в новую копию.")
                .foregroundStyle(.secondary)
                .font(.callout)
        }
        .frame(maxWidth: .infinity)
        .padding(.vertical, 22)
        .background(isTargeted ? Color.accentColor.opacity(0.14) : Color.secondary.opacity(0.08))
        .clipShape(RoundedRectangle(cornerRadius: 8))
        .overlay(
            RoundedRectangle(cornerRadius: 8)
                .strokeBorder(isTargeted ? Color.accentColor : Color.secondary.opacity(0.2), style: StrokeStyle(lineWidth: 1, dash: [6]))
        )
    }
}
