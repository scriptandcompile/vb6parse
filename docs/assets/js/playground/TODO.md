# VB6Parse Playground - Implementation TODO Checklist

This document outlines all the tasks needed to complete the playground implementation.
The files have been created with placeholder/skeleton code - you need to fill in the actual implementations.

## 📋 Phase 1: WASM Integration (Critical - Do First!)

### WASM Module (`src/wasm.rs` or `src/playground.rs`)
- [ ] Create new Rust module for WASM bindings
- [ ] Add wasm-bindgen functions:
  - [ ] `init_panic_hook()` - Set up panic handler
  - [ ] `parse_vb6_code(code: &str, file_type: &str) -> Result<JsValue, JsValue>`
  - [ ] `tokenize_vb6_code(code: &str) -> Result<JsValue, JsValue>`
- [ ] Create serializable output structures:
  - [ ] `PlaygroundOutput` with tokens, CST, errors, stats
  - [ ] `TokenInfo` structure
  - [ ] `CstNodeJson` structure (serializable version of CstNode)
  - [ ] `ErrorInfo` structure
- [ ] Implement CST to JSON conversion
- [ ] Add proper error handling and Result types
- [ ] Test WASM functions in Rust with unit tests

### Cargo.toml Configuration
- [ ] Add `crate-type = ["cdylib", "rlib"]` to `[lib]`
- [ ] Add `serde_json` dependency (if not present)
- [ ] Add WASM target dependencies under `[target.'cfg(target_arch = "wasm32")'.dependencies]`
- [ ] Verify build profile optimizations

### Build Script
- [ ] Test `scripts/build-wasm.py` on your platform
- [ ] Verify wasm-pack is installed: `cargo install wasm-pack`
- [ ] Verify wasm-opt is installed: `cargo install wasm-opt`
- [ ] Run build: `python scripts/build-wasm.py --optimize`
- [ ] Verify output in `docs/assets/wasm/`:
  - [ ] `vb6parse_bg.wasm`
  - [ ] `vb6parse.js`
  - [ ] `vb6parse_bg.wasm.d.ts` (if TypeScript not disabled)

## 📋 Phase 2: JavaScript Integration

### Parser Module (`parser.js`)
- [ ] Replace mock `initWasm()` with actual WASM initialization
  - [ ] Import from `../wasm/vb6parse.js`
  - [ ] Call `init()` function
  - [ ] Call `init_panic_hook()`
- [ ] Replace mock `parseCode()` with actual WASM call
  - [ ] Call `parse_vb6_code(code, fileType)`
  - [ ] Convert JsValue to JavaScript object
  - [ ] Add proper error handling
- [ ] Replace mock `tokenizeCode()` with actual WASM call
  - [ ] Call `tokenize_vb6_code(code)`
  - [ ] Handle returned tokens
- [ ] Remove all mock data functions:
  - [ ] `createMockParseResult()`
  - [ ] `createMockTokens()`
  - [ ] `createMockCst()`
- [ ] Test WASM integration with real VB6 code

### Editor Module (`editor.js`)
- [ ] Complete VB6 language definition for Monaco:
  - [ ] Add all VB6 keywords (160+ from library)
  - [ ] Add all VB6 operators
  - [ ] Improve string and comment tokenization
  - [ ] Add support for line continuations (`_`)
  - [ ] Add support for `Attribute` statements
- [ ] Implement `highlightRange()` function:
  - [ ] Use `editor.deltaDecorations()`
  - [ ] Add highlight styling
  - [ ] Clear previous highlights
- [ ] Implement localStorage persistence:
  - [ ] Complete `saveToLocalStorage()`
  - [ ] Complete `loadFromLocalStorage()`
  - [ ] Save on every change (debounced)
- [ ] Test editor with various VB6 code samples
- [ ] Optional: Consider switching to Ace Editor if Monaco is too heavy

### Renderer Module (`renderer.js`)
- [ ] Test token rendering with real parse results
- [ ] Implement token filtering by type
- [ ] Implement token search functionality
- [ ] Complete `highlightTokenInEditor()`:
  - [ ] Import Editor module
  - [ ] Call `Editor.highlightRange()`
- [ ] Complete `highlightNodeInEditor()`:
  - [ ] Convert byte range to line/column
  - [ ] Highlight in editor
- [ ] Test CST rendering with real parse results
- [ ] Test Info tab with real statistics
- [ ] Verify error/warning display

### Tree Visualization Module (`tree-viz.js`)
- [ ] Implement D3 tree layout:
  - [ ] Complete `convertCstToD3Format()`
  - [ ] Configure `d3.tree()` layout
  - [ ] Implement vertical layout
  - [ ] Implement horizontal layout
- [ ] Implement node rendering:
  - [ ] Use D3 data binding
  - [ ] Draw circles with correct colors
  - [ ] Add text labels
  - [ ] Implement hover effects
  - [ ] Add click handlers
- [ ] Implement link rendering:
  - [ ] Use `d3.linkVertical()` and `d3.linkHorizontal()`
  - [ ] Draw curved paths
  - [ ] Add animations
- [ ] Implement interactions:
  - [ ] Node selection
  - [ ] Tooltip on hover
  - [ ] Details panel
  - [ ] Highlight in editor on node click
- [ ] Complete `getTreeStats()` for accurate counts
- [ ] Test with small and large trees
- [ ] Optimize for performance (>500 nodes)

### Main Module (`main.js`)
- [ ] Test all event handlers
- [ ] Implement share functionality:
  - [ ] Encode code with `btoa()` or LZ-string
  - [ ] Generate shareable URL
  - [ ] Copy to clipboard
  - [ ] Show share dialog
- [ ] Implement URL parameter loading:
  - [ ] Parse `?code=` and `?type=` parameters
  - [ ] Decode and load code
  - [ ] Set file type
- [ ] Test localStorage save/load
- [ ] Test debounced auto-parse
- [ ] Verify all keyboard interactions work
- [ ] Test split panel resizer thoroughly

## 📋 Phase 3: UI/UX Polish

### HTML (`playground.html`)
- [ ] Test on multiple browsers (Chrome, Firefox, Safari, Edge)
- [ ] Verify all IDs and classes match JS/CSS
- [ ] Test responsive design on mobile/tablet
- [ ] Add meta tags for social sharing
- [ ] Test accessibility with screen reader
- [ ] Verify ARIA labels

### CSS (`playground.css` & `tree-viz.css`)
- [ ] Test light and dark themes
- [ ] Verify all color variables work correctly
- [ ] Test responsive breakpoints
- [ ] Verify print styles
- [ ] Test on different screen sizes
- [ ] Check high contrast mode
- [ ] Verify reduced motion support

### Examples (`examples.js`)
- [ ] Add more realistic VB6 examples
- [ ] Include examples with errors for testing
- [ ] Add complex examples (databases, APIs, etc.)
- [ ] Add descriptions for each example
- [ ] Consider loading from external JSON file
- [ ] Test all examples load and parse correctly

## 📋 Phase 4: Integration with Main Site

### Navigation
- [ ] Update `docs/index.html` to add Playground link
- [ ] Update navigation in all other docs pages
- [ ] Add Playground section to README
- [ ] Update documentation to reference playground
- [ ] Add screenshots/GIFs to docs

### GitHub Pages Deployment
- [ ] Test locally with `python -m http.server` in `docs/`
- [ ] Verify WASM files are included in git (or generated by CI)
- [ ] Test deployment on GitHub Pages
- [ ] Verify all assets load correctly (no CORS issues)
- [ ] Test with custom domain if applicable

### CI/CD (Optional)
- [ ] Create `.github/workflows/build-wasm.yml`
- [ ] Test workflow on all platforms
- [ ] Set up automatic WASM builds on push
- [ ] Configure caching for faster builds
- [ ] Add workflow status badge to README

## 📋 Phase 5: Testing & Quality Assurance

### Functional Testing
- [ ] Test with empty code
- [ ] Test with invalid VB6 syntax
- [ ] Test with very long files (>10,000 lines)
- [ ] Test with Unicode/special characters
- [ ] Test all file types (module, class, form, project)
- [ ] Test error handling (WASM panics, etc.)
- [ ] Test browser back/forward buttons
- [ ] Test page refresh (state preservation)

### Performance Testing
- [ ] Measure parse time for various file sizes:
  - [ ] Small (<100 lines)
  - [ ] Medium (100-1000 lines)
  - [ ] Large (1000-10000 lines)
  - [ ] Very large (>10000 lines)
- [ ] Profile WASM memory usage
- [ ] Optimize bundle size (target <2MB total)
- [ ] Test tree rendering performance
- [ ] Consider Web Workers for large files

### Cross-browser Testing
- [ ] Chrome (latest)
- [ ] Firefox (latest)
- [ ] Safari (latest)
- [ ] Edge (latest)
- [ ] Mobile Safari (iOS)
- [ ] Mobile Chrome (Android)

### Accessibility Testing
- [ ] Keyboard navigation works throughout
- [ ] Screen reader announces all actions
- [ ] Focus management is logical
- [ ] Color contrast meets WCAG AA
- [ ] No motion for users with reduced motion preference

## 📋 Phase 6: Documentation

### User Documentation
- [ ] Create `docs/playground-guide.html`
- [ ] Add tutorial for first-time users
- [ ] Document all features
- [ ] Add keyboard shortcuts reference
- [ ] Include troubleshooting section
- [ ] Add FAQ

### Developer Documentation
- [ ] Document WASM module API
- [ ] Add architecture diagram
- [ ] Document JavaScript module interactions
- [ ] Add contribution guidelines for playground
- [ ] Document build and deployment process

## 📋 Phase 7: Optional Enhancements

### Features
- [ ] File upload (drag & drop)
- [ ] Download parse results as JSON
- [ ] Export tree as SVG/PNG
- [ ] Multiple file tabs
- [ ] Parse history
- [ ] Code snippets library
- [ ] Syntax error highlighting in editor
- [ ] IntelliSense/autocomplete
- [ ] Find in output
- [ ] Diff view (compare two versions)

### Advanced
- [ ] Web Workers for parsing
- [ ] Progressive Web App (PWA)
- [ ] Offline support
- [ ] VS Code extension using same WASM
- [ ] API endpoint for parsing (if desired)
- [ ] Collaboration features
- [ ] Analytics integration

## 🎯 Quick Start Guide

To get started implementing:

1. **Start with WASM** - Nothing else will work without it:
   ```bash
   # Create src/wasm.rs with bindings
   # Update Cargo.toml
   # Build: python scripts/build-wasm.py --optimize
   ```

2. **Test WASM in browser console**:
   ```javascript
   import init, { parse_vb6_code } from './assets/wasm/vb6parse.js';
   await init();
   const result = parse_vb6_code('Option Explicit\\n\\nSub Test()\\nEnd Sub', 'module');
   console.log(result);
   ```

3. **Update parser.js** to use real WASM (remove all mocks)

4. **Test with simple VB6 code** - verify tokens and CST render

5. **Implement tree visualization** with D3.js

6. **Polish UI/UX** and add remaining features

7. **Test thoroughly** on all browsers

8. **Deploy to GitHub Pages**

## 📞 Getting Help

- Check VB6Parse documentation: `cargo doc --open`
- Review wasm-bindgen docs: https://rustwasm.github.io/wasm-bindgen/
- Review Monaco Editor docs: https://microsoft.github.io/monaco-editor/
- Review D3.js docs: https://d3js.org/
- Look at similar playgrounds for inspiration:
  - Rust Playground: https://play.rust-lang.org/
  - AST Explorer: https://astexplorer.net/

## ✅ Definition of Done

The playground is complete when:

- [ ] WASM module builds and loads successfully
- [ ] All four output tabs render correctly with real data
- [ ] Tree visualization shows interactive D3 tree
- [ ] All examples load and parse
- [ ] Error handling works properly
- [ ] Mobile/responsive design works
- [ ] Works on all major browsers
- [ ] Integrated into main documentation site
- [ ] User documentation is complete
- [ ] No console errors or warnings

Good luck! 🚀
