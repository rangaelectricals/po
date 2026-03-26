# UI & UX Instructions and Design Patterns

## 1. Layout & Structure
- Responsive layout: Adapt to desktop, tablet, and mobile.
- Sidebar navigation: Collapsible sidebar (drawer) for desktop, slide-out menu for mobile.
- Top bar: Fixed for branding, user profile, and quick actions (notifications, logout).

## 2. Navigation
- Highlight current page in sidebar.
- Use breadcrumbs on detail/sub-pages (e.g., PO View, Edit).
- Keep navigation consistent across all pages.

## 3. Forms & Data Entry
- Clear labels and helper text for all fields.
- Group related fields with section headers.
- Input validation with inline error display.
- Show loading indicators on submit actions.

## 4. Tables & Lists
- Striped rows and hover effects for tables (e.g., PO List, Vendor Master).
- Allow sorting and filtering on key columns.
- Paginate long lists with clear navigation controls.

## 5. Modals & Dialogs
- Use modals for confirmation, editing, or adding new items.
- Always provide a clear way to close modals (X button and outside click).

## 6. Buttons & Actions
- Primary color for main actions (e.g., Save, Add PO).
- Secondary/outline buttons for less important actions (e.g., Cancel).
- Disable buttons when actions are not available.

## 7. Feedback & Notifications
- Toast notifications for success, error, and info messages.
- Inline alerts for form errors or important warnings.

## 8. Accessibility
- All interactive elements are keyboard accessible.
- Sufficient color contrast for text and UI elements.
- Use aria-labels and roles where appropriate.

## 9. Visual Design
- Modern, clean design: ample whitespace, clear typography, limited color palette.
- Icons alongside text for navigation and actions.
- Consistent spacing, font sizes, and button styles.

## 10. Patterns to Follow
- Master-detail: PO list → PO view/edit.
- CRUD: Consistent create, read, update, delete flows for all entities.
- Search/filter: Prominently placed search bars for lists.
- Profile/settings: User profile and settings accessible from the top bar.

## Reference Modern UI Libraries
- Use design inspiration from Material Design, Ant Design, or Bootstrap for component styles and interactions.

---

If you need a specific page or component design, specify which one for a detailed wireframe or code example.
