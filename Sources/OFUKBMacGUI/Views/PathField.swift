import SwiftUI

struct PathField: View {
    let title: String
    let value: String
    let placeholder: String
    let systemImage: String
    let buttonTitle: String
    let action: () -> Void

    var body: some View {
        HStack(spacing: 12) {
            Label(title, systemImage: systemImage)
                .frame(width: 130, alignment: .leading)
                .foregroundStyle(.secondary)

            Text(value.isEmpty ? placeholder : value)
                .lineLimit(1)
                .truncationMode(.middle)
                .foregroundStyle(value.isEmpty ? .secondary : .primary)
                .frame(maxWidth: .infinity, alignment: .leading)
                .padding(.horizontal, 10)
                .padding(.vertical, 7)
                .background(Color.secondary.opacity(0.08))
                .clipShape(RoundedRectangle(cornerRadius: 6))

            Button(buttonTitle, action: action)
        }
    }
}
