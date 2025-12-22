# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is a **SharePoint Framework (SPFx) 1.22.1** web part project called "SmartForm". It targets Microsoft 365, SharePoint, and Microsoft Teams environments, using the Heft build system from Rush Stack.

## Build System & Commands

This project uses **Heft** (not gulp) as its build system. All commands must use `heft`, not npm scripts like `npm run build`.

### Development Commands

- `heft start --clean` - Start development server with live reload (serves on https://localhost:4321)
- `heft build` - Build the project
- `heft test` - Run tests
- `heft clean` - Clean build artifacts
- `heft build --clean --production` - Production build
- `heft package-solution --production` - Create .sppkg package for deployment

### Testing

- `heft test` - Run all tests
- `heft test-watch` - Run tests in watch mode

### Available Phases

Heft supports these phases (use `heft <phase>` or `heft <phase>-watch`):
- `build`, `test`, `package-solution`, `dev-deploy`, `deploy-azure-storage`, `eject-webpack`

## Architecture

### Web Part Structure

SPFx web parts follow a specific architecture pattern:

1. **Web Part Class** (`SmartFormWebPart.ts`):
   - Extends `BaseClientSideWebPart<TProps>`
   - Handles lifecycle: `onInit()`, `render()`, `onDispose()`
   - Manages theme changes via `onThemeChanged()`
   - Detects environment (SharePoint/Teams/Outlook/Office) in `_getEnvironmentMessage()`
   - Provides property pane configuration via `getPropertyPaneConfiguration()`

2. **React Component** (`SmartForm.tsx`):
   - Receives props defined in `ISmartFormProps` interface
   - Props include: `description`, `isDarkTheme`, `environmentMessage`, `hasTeamsContext`, `userDisplayName`
   - Uses CSS Modules for styling (`SmartForm.module.scss`)

3. **Props Interface** (`ISmartFormProps.ts`):
   - Defines the contract between web part and React component
   - Web part props (`ISmartFormWebPartProps`) are separate from component props

### Key SPFx Concepts

- **Context**: `this.context` provides access to page context, Teams SDK, user info, etc.
- **Property Pane**: Configured in `getPropertyPaneConfiguration()`, allows users to configure the web part
- **Theme Support**: Dark mode handled via `isDarkTheme` prop and CSS variables
- **Multi-Environment**: Code detects if running in SharePoint, Teams, Outlook, or Office
- **Localization**: String resources in `src/webparts/smartForm/loc/`

### Directory Structure

```
src/webparts/smartForm/
├── SmartFormWebPart.ts          # Web part entry point
├── SmartFormWebPart.manifest.json
├── components/
│   ├── SmartForm.tsx            # React component
│   ├── SmartForm.module.scss    # CSS Modules styles
│   └── ISmartFormProps.ts       # Component props interface
├── loc/                         # Localization resources
└── assets/                      # Images and static assets
```

### Configuration Files

- `config/package-solution.json` - Solution metadata and package configuration
- `config/serve.json` - Dev server settings (port, https, initial page)
- `.yo-rc.json` - Yeoman generator configuration
- TypeScript config extends from `@microsoft/spfx-web-build-rig`

## Dependencies

### Core SPFx Libraries
- `@microsoft/sp-webpart-base` - Web part base classes
- `@microsoft/sp-core-library` - Core utilities
- `@microsoft/sp-property-pane` - Property pane controls
- `@microsoft/sp-component-base` - Component base classes

### UI Framework
- React 17.0.1
- `@fluentui/react` ^8.106.4 (Fluent UI for React)

### Node Version
- Requires Node.js >=22.14.0 <23.0.0

## Important Patterns

### Rendering Pattern
```typescript
public render(): void {
  const element = React.createElement(Component, props);
  ReactDom.render(element, this.domElement);
}
```

### Escaping User Input
Always escape user-provided content using `escape()` from `@microsoft/sp-lodash-subset` to prevent XSS.

### Theme CSS Variables
The web part sets CSS custom properties for theming:
- `--bodyText`
- `--link`
- `--linkHovered`

These are set in `onThemeChanged()` and can be used in SCSS files.

## Deployment

The built package is created at `solution/smart-form.sppkg` (configured in `config/package-solution.json`). This .sppkg file is deployed to the SharePoint App Catalog.
