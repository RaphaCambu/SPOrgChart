# `sporgchart`

Standalone React organizational chart component extracted from the SPFx dynamic JSX workspace.

## Install

```bash
npm install sporgchart
```

Also install peer dependencies if your project does not already include them:

```bash
npm install react react-dom @fluentui/react
```

## Usage

```tsx
import * as React from 'react';
import { SPOrgChart } from 'sporgchart';
import 'sporgchart/styles.css';

export function Example(): JSX.Element {
  return (
    <SPOrgChart
      items={users}
      rootPersonId="ceo-id"
      expandAll={true}
      idField="ID"
      parentIdField="ManagerId"
      nameField="DisplayName"
      positionField="JobTitle"
      photoField="PhotoUrl"
      emailField="EMail"
      linkField="ProfileUrl"
    />
  );
}
```

## Notes

- The package name should be published in lowercase on npm.
- The component expects a flat array of people with an id field and a parent/manager id field.
- Import `styles.css` to get the org-chart presentation layer.
