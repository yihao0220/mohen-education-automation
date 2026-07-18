import AppKit
import Foundation
import WebKit

guard CommandLine.arguments.count == 3 else {
    fputs("usage: quicklook_webkit_pdf INPUT.html OUTPUT.pdf\n", stderr)
    exit(2)
}

let inputURL = URL(fileURLWithPath: CommandLine.arguments[1]).standardizedFileURL
let outputURL = URL(fileURLWithPath: CommandLine.arguments[2]).standardizedFileURL
let application = NSApplication.shared
let webView = WKWebView(frame: NSRect(x: 0, y: 0, width: 595, height: 842))
let window = NSWindow(
    contentRect: NSRect(x: 0, y: 0, width: 595, height: 842),
    styleMask: [.borderless],
    backing: .buffered,
    defer: false
)
window.contentView = webView
window.orderOut(nil)

final class Renderer: NSObject, WKNavigationDelegate {
    let webView: WKWebView
    let outputURL: URL

    init(webView: WKWebView, outputURL: URL) {
        self.webView = webView
        self.outputURL = outputURL
    }

    func webView(_ webView: WKWebView, didFinish navigation: WKNavigation!) {
        DispatchQueue.main.asyncAfter(deadline: .now() + 0.5) {
            self.webView.evaluateJavaScript(
                "Math.max(document.body.scrollHeight, document.documentElement.scrollHeight)"
            ) { value, error in
                if let error = error {
                    fputs("WebKit height query failed: \(error)\n", stderr)
                    exit(1)
                }
                guard let height = value as? Double, height > 0 else {
                    fputs("WebKit returned invalid document height\n", stderr)
                    exit(1)
                }
                let configuration = WKPDFConfiguration()
                configuration.rect = CGRect(x: 0, y: 0, width: 595, height: height)
                self.webView.createPDF(configuration: configuration) { result in
                    do {
                        let data = try result.get()
                        try data.write(to: self.outputURL, options: .atomic)
                        exit(0)
                    } catch {
                        fputs("WebKit PDF creation failed: \(error)\n", stderr)
                        exit(1)
                    }
                }
            }
        }
    }

    func webView(_ webView: WKWebView, didFail navigation: WKNavigation!, withError error: Error) {
        fputs("WebKit navigation failed: \(error)\n", stderr)
        exit(1)
    }
}

let renderer = Renderer(webView: webView, outputURL: outputURL)
webView.navigationDelegate = renderer
webView.loadFileURL(inputURL, allowingReadAccessTo: inputURL.deletingLastPathComponent())
application.run()
