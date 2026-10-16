# AI Office removal map

The editor and Agents windows no longer import the developer feature entry points listed below, in either desktop or browser builds. Word's custom editor, webviews, chat, MCP, comments, terminal and multi-file diff review remain available.

| Feature | Removed from the application | Retained dependency |
|---|---|---|
| Notebooks | Notebook editor/diff UI, toolbar, settings and notebook contribution entry point; built-in IPYNB serializer and output renderers | Notebook API services used by the extension host, bulk edits and chat; registrations are separated into `notebook.service.contribution.ts` |
| Debug | Run and Debug views, breakpoints, debug toolbar, launch commands and built-in JS debuggers | Upstream null debug/visualizer services, plus the platform extension-host debug service |
| Testing | Test Explorer, results/coverage UI, test commands and testing agent tools | Upstream null test, profile and result services required by extension-host and chat consumers |
| Tasks | Run/build task commands, task configuration and task UI contribution; built-in npm/Grunt/Gulp/Jake task extensions | Empty task API service; no task discovery, execution, commands or watchers; the terminal remains available to agents |
| Merge editor | Three-way merge editor and its desktop commands; built-in merge-conflict extension | Text diff and multi-file diff review; shared model code used by other editors |
| Issue reporter | Reporter UI, report-issue commands, screenshots and recording contributions | A small `IWorkbenchIssueService` implementation that reports the feature as unavailable to extension management/Settings Sync |
| Remote | Remote Explorer, status indicator, Ports view/forwarding UI, browser remote-start entry and remote tunnel-host UI | Remote filesystem, extension-host and agent transport services needed for server-hosted operation |
| Snippets | Snippet configuration/insertion commands, tab completion, user snippets and snippet suggestions | Empty snippets API service; Monaco's low-level snippet insertion used by chat remains available |
| Interactive | Interactive Window and REPL notebook contributions | Interactive document/history APIs required by extension-host participants |

## Extension exclusion

`product.json` lists excluded built-ins in `disabledBuiltinExtensions`. Desktop and browser scanners honor this list, including development checkouts. Local extension packaging also excludes these extensions; their compile/media entries and downloaded JavaScript debugger definitions have been removed.

Of the 98 built-in manifests in this checkout, 64 are excluded and 34 retained. This includes 54 newly excluded extensions beyond the initial feature cleanup. The earlier estimate of 106 extensions was not the inventory of this branch.

| Retained group | Extensions and purpose |
|---|---|
| Office documents | `ai-office-pdf` (Word and PDF), `ai-office-slides` |
| Documents and previews | Markdown language/editor, math, Mermaid, media preview, simple browser |
| Agent access | Copilot, prompt-file support, GitHub and Microsoft authentication |
| History | Git, Git Base, GitHub integration, diff grammar |
| Configuration and diagnostics | Configuration editing, JSON grammar/language service, YAML for agent frontmatter, logs and search results |
| Appearance | Existing color and icon themes, including the AI Office icon theme |

Excluded extensions include Python, Java, Rust, C/C++, C#, F#, JavaScript/TypeScript, Go, Ruby, PHP, Docker, shell and other programming grammars; CSS/HTML/PHP/TypeScript language services; Emmet, extension-authoring, references and terminal suggestions; development test extensions; and the notebook, debugging, task, merge-conflict and tunnel extensions excluded earlier. Configuration and document support stays available without programming-language intelligence.

`build/lib/builtinExtensionPolicy.ts` applies the same product policy to packaging, TypeScript compilation, browser extension builds and media builds. Browser build discovery also filters excluded manifests, so stale generated extension output cannot reintroduce them into the browser's bundled extension list. Source directories and dependency lockfiles remain for upstream maintenance and repository development.

The upstream source directories remain in the repository because shared code and extension-host API implementations still import them. Removing an entry point removes its feature registrations, not every transitive model/helper from the application. The earlier approximate line counts are not a measured bundle-size reduction. Further source deletion requires separating those shared APIs first.

## Validation

After installing root and build dependencies:

```sh
node --test build/lib/test/officeWorkbench.test.ts
node --test build/lib/test/builtinExtensionPolicy.test.ts
```

This bundles the actual desktop/browser entry points for both windows in memory and checks that excluded feature entry points are absent while document and agent infrastructure remain reachable. It does not launch Electron or a browser.

The extension-scanner suite includes a regression test that excludes a product-disabled built-in from fresh and cached scans while keeping the Word/PDF extension. Run it with the existing VS Code unit-test runner after compiling the client.

The built-in policy tests verify the retained inventory, its declared extension dependencies, and policy selection for relative and absolute desktop/browser/media build inputs.
