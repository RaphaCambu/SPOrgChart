import * as React from 'react';
import {
  DefaultButton,
  Dialog,
  IconButton,
  Link,
  Persona,
  PersonaSize,
  Stack,
  Text
} from '@fluentui/react';

interface IOrgChartNode {
  item: any;
  id: string;
  parentId?: string;
  children: IOrgChartNode[];
}

export interface ISPOrgChartProps {
  items?: any[];
  rootPersonId?: string | number;
  rootPosition?: string;
  rootEmail?: string;
  rootLoginName?: string;
  idField?: string;
  parentIdField?: string;
  nameField?: string;
  positionField?: string;
  photoField?: string;
  emailField?: string;
  loginNameField?: string;
  linkField?: string;
  initiallyExpanded?: boolean;
  expandAll?: boolean;
  indentSize?: number;
  emptyMessage?: string;
}

const DEFAULT_PROPS = {
  idField: 'ID',
  parentIdField: 'ManagerId',
  nameField: 'DisplayName',
  positionField: 'JobTitle',
  photoField: 'PhotoUrl',
  emailField: 'EMail',
  loginNameField: 'UserPrincipalName',
  linkField: 'ProfileUrl',
  initiallyExpanded: true,
  expandAll: true,
  indentSize: 24,
  emptyMessage: 'No organization data available.'
};

const normalizeValue = (value: any): string | undefined => {
  if (value === null || value === undefined) {
    return undefined;
  }

  if (typeof value === 'object') {
    const nestedCandidate =
      value.Id ??
      value.ID ??
      value.id ??
      value.Value ??
      value.value ??
      value.Email ??
      value.EMail ??
      value.email ??
      value.LoginName ??
      value.loginName ??
      value.UserPrincipalName ??
      value.userPrincipalName ??
      value.DisplayName ??
      value.displayName ??
      value.Title ??
      value.title;

    if (nestedCandidate !== undefined && nestedCandidate !== null) {
      return String(nestedCandidate);
    }
  }

  return String(value);
};

const getFieldValue = (item: any, fieldName?: string): any => {
  if (!item || !fieldName) {
    return undefined;
  }

  if (Object.prototype.hasOwnProperty.call(item, fieldName)) {
    return item[fieldName];
  }

  const pathSegments = fieldName.split(/[/.]/).filter(Boolean);

  if (!pathSegments.length) {
    return undefined;
  }

  return pathSegments.reduce((current: any, segment: string) => {
    if (current === null || current === undefined) {
      return undefined;
    }

    return current[segment];
  }, item);
};

const materializeItems = (items: any, fieldNames: string[]): any[] => {
  if (!items) {
    return [];
  }

  const length = typeof items.length === 'number' ? items.length : 0;
  const materialized: any[] = [];

  for (let index = 0; index < length; index += 1) {
    const item = items[index];

    if (!item) {
      continue;
    }

    fieldNames.forEach((fieldName: string) => {
      if (!fieldName) {
        return;
      }

      // Access configured fields eagerly so lazy SPData collections resolve before tree construction.
      void getFieldValue(item, fieldName);
    });

    materialized.push(item);
  }

  return materialized;
};

const buildTeamsChatUrl = (email?: string): string => {
  if (!email) {
    return '';
  }

  return `https://teams.microsoft.com/l/chat/0/0?users=${encodeURIComponent(email)}`;
};

const normalizePresence = (value: any): string => {
  return String(value || '').toLowerCase();
};

const getPresenceMeta = (availability: any): { label: string; className: string } => {
  switch (normalizePresence(availability)) {
    case 'available':
      return { label: 'Available', className: 'sp-orgchart-presenceAvailable' };
    case 'busy':
      return { label: 'Busy', className: 'sp-orgchart-presenceBusy' };
    case 'donotdisturb':
      return { label: 'Do not disturb', className: 'sp-orgchart-presenceDnd' };
    case 'away':
    case 'berightback':
      return { label: 'Away', className: 'sp-orgchart-presenceAway' };
    case 'offline':
    case 'presenceunknown':
      return { label: 'Offline', className: 'sp-orgchart-presenceOffline' };
    default:
      return { label: 'Status unavailable', className: 'sp-orgchart-presenceOffline' };
  }
};

const buildTree = (
  items: any[],
  idField: string,
  parentIdField: string
): { nodes: Map<string, IOrgChartNode>; roots: IOrgChartNode[] } => {
  const nodes = new Map<string, IOrgChartNode>();
  const roots: IOrgChartNode[] = [];

  items.forEach((item: any) => {
    const id = normalizeValue(getFieldValue(item, idField));

    if (!id) {
      return;
    }

    const parentId = normalizeValue(getFieldValue(item, parentIdField));

    nodes.set(id, {
      item,
      id,
      parentId,
      children: []
    });
  });

  nodes.forEach((node: IOrgChartNode) => {
    if (node.parentId && nodes.has(node.parentId)) {
      nodes.get(node.parentId)!.children.push(node);
      return;
    }

    roots.push(node);
  });

  return { nodes, roots };
};

const findRootNode = (
  nodes: Map<string, IOrgChartNode>,
  roots: IOrgChartNode[],
  props: Required<ISPOrgChartProps>
): IOrgChartNode | undefined => {
  if (props.rootPersonId !== undefined && props.rootPersonId !== null && props.rootPersonId !== '') {
    return nodes.get(String(props.rootPersonId));
  }

  const allNodes = Array.from(nodes.values());

  if (props.rootEmail) {
    return allNodes.find((node: IOrgChartNode) => {
      const value = normalizeValue(getFieldValue(node.item, props.emailField));
      return value?.toLowerCase() === props.rootEmail.toLowerCase();
    });
  }

  if (props.rootLoginName) {
    return allNodes.find((node: IOrgChartNode) => {
      const value = normalizeValue(getFieldValue(node.item, props.loginNameField));
      return value?.toLowerCase() === props.rootLoginName.toLowerCase();
    });
  }

  if (props.rootPosition) {
    return allNodes.find((node: IOrgChartNode) => {
      const value = normalizeValue(getFieldValue(node.item, props.positionField));
      return value?.toLowerCase() === props.rootPosition.toLowerCase();
    });
  }

  return roots[0];
};

const collectExpandableNodes = (node: IOrgChartNode, result: Record<string, boolean>): void => {
  if (node.children.length > 0) {
    result[node.id] = true;
  }

  node.children.forEach((child: IOrgChartNode) => collectExpandableNodes(child, result));
};

const cardContainerStyles: React.CSSProperties = {
  border: '1px solid #d1d1d1',
  borderRadius: 10,
  padding: 12,
  backgroundColor: '#ffffff',
  boxShadow: '0 1px 2px rgba(0, 0, 0, 0.08)',
  minWidth: 240,
  maxWidth: 360
};

const toggleButtonStyles = {
  root: {
    width: 34,
    height: 34,
    borderRadius: '50%',
    border: '1px solid #c8c8c8',
    backgroundColor: '#ffffff'
  }
};

const OrgChartNodeView: React.FC<{
  node: IOrgChartNode;
  level: number;
  expandedKeys: Record<string, boolean>;
  onToggle: (nodeId: string) => void;
  resolvedProps: Required<ISPOrgChartProps>;
}> = ({ node, level, expandedKeys, onToggle, resolvedProps }) => {
  const hasChildren = node.children.length > 0;
  const isExpanded = expandedKeys[node.id] !== false;
  const useVerticalLeafBranch =
    hasChildren && node.children.every((child: IOrgChartNode) => child.children.length === 0);
  const name = normalizeValue(getFieldValue(node.item, resolvedProps.nameField)) || 'Unnamed person';
  const position = normalizeValue(getFieldValue(node.item, resolvedProps.positionField)) || '';
  const photoUrl = normalizeValue(getFieldValue(node.item, resolvedProps.photoField));
  const email = normalizeValue(getFieldValue(node.item, resolvedProps.emailField));
  const linkUrl = normalizeValue(getFieldValue(node.item, resolvedProps.linkField));
  const manager = getFieldValue(node.item, 'Manager');
  const managerName =
    normalizeValue(getFieldValue(manager, resolvedProps.nameField)) ||
    normalizeValue(getFieldValue(manager, 'DisplayName')) ||
    '';
  const department = normalizeValue(getFieldValue(node.item, 'Department')) || '';
  const officeLocation = normalizeValue(getFieldValue(node.item, 'OfficeLocation')) || '';
  const presenceMeta = getPresenceMeta(getFieldValue(node.item, 'Availability'));
  const teamsChatUrl = buildTeamsChatUrl(email || undefined);
  const [isDialogOpen, setIsDialogOpen] = React.useState<boolean>(false);

  const cardContent = (
    <>
      <Stack
        horizontalAlign="stretch"
        tokens={{ childrenGap: 10 }}
        style={cardContainerStyles}
        className={`sp-orgchart-card sp-orgchart-cardLevel${level === 0 ? '0' : level === 1 ? '1' : '2'}`}
        role="treeitem"
        aria-expanded={hasChildren ? isExpanded : undefined}
        aria-level={level + 1}
      >
        <Stack horizontalAlign="center" className="sp-orgchart-avatarWrap">
          {photoUrl ? (
            <img src={photoUrl} alt={name} className="sp-orgchart-avatar" />
          ) : (
            <Stack horizontalAlign="center" verticalAlign="center" className="sp-orgchart-avatarFallback">
              <Text>{name.charAt(0)}</Text>
            </Stack>
          )}
        </Stack>

        <Stack className="sp-orgchart-titleBlock" tokens={{ childrenGap: 4 }}>
          <IconButton
            iconProps={{ iconName: 'Info' }}
            title="Open details"
            ariaLabel={`Open details for ${name}`}
            className="sp-orgchart-moreButton"
            onClick={(event?: React.MouseEvent<any>) => {
              event?.preventDefault();
              event?.stopPropagation();
              setIsDialogOpen(true);
            }}
          />
          <Text variant="medium" className="sp-orgchart-nameText">
            {name}
          </Text>
          {position ? (
            <Text variant="small" className="sp-orgchart-roleText">
              {position}
            </Text>
          ) : null}

          <Stack horizontal verticalAlign="center" horizontalAlign="space-between" className="sp-orgchart-footerBar">
            <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
              <span className={`sp-orgchart-presenceDot ${presenceMeta.className}`} />
              <Text variant="xSmall" className="sp-orgchart-presenceLabel">
                {presenceMeta.label}
              </Text>
            </Stack>
            {teamsChatUrl ? (
              <DefaultButton
                text="Chat"
                iconProps={{ iconName: 'TeamsLogo' }}
                href={teamsChatUrl}
                target="_blank"
                className="sp-orgchart-chatButton"
              />
            ) : null}
          </Stack>
        </Stack>
      </Stack>

      <Dialog
        hidden={!isDialogOpen}
        onDismiss={() => setIsDialogOpen(false)}
        dialogContentProps={{
          title: ''
        }}
      >
        <Stack tokens={{ childrenGap: 14 }}>
          <Persona
            text={name}
            secondaryText={position}
            imageUrl={photoUrl}
            size={PersonaSize.size72}
          />
          {email ? (
            <Stack tokens={{ childrenGap: 4 }}>
              <Text variant="mediumPlus">Email</Text>
              <Link href={`mailto:${email}`}>{email}</Link>
            </Stack>
          ) : null}
          {managerName ? (
            <Stack tokens={{ childrenGap: 4 }}>
              <Text variant="mediumPlus">Manager</Text>
              <Text>{managerName}</Text>
            </Stack>
          ) : null}
          {department ? (
            <Stack tokens={{ childrenGap: 4 }}>
              <Text variant="mediumPlus">Department</Text>
              <Text>{department}</Text>
            </Stack>
          ) : null}
          {officeLocation ? (
            <Stack tokens={{ childrenGap: 4 }}>
              <Text variant="mediumPlus">Office</Text>
              <Text>{officeLocation}</Text>
            </Stack>
          ) : null}
        </Stack>
      </Dialog>
    </>
  );

  return (
    <li className="sp-orgchart-nodeItem">
      <div className="sp-orgchart-nodeShell">
        {linkUrl ? (
          <Link href={linkUrl} target="_blank" underline className="sp-orgchart-cardLink">
            {cardContent}
          </Link>
        ) : (
          cardContent
        )}

        {hasChildren ? (
          <IconButton
            iconProps={{ iconName: isExpanded ? 'Remove' : 'Add' }}
            title={isExpanded ? 'Collapse' : 'Expand'}
            ariaLabel={isExpanded ? `Collapse ${name}` : `Expand ${name}`}
            styles={toggleButtonStyles}
            onClick={() => onToggle(node.id)}
            className="sp-orgchart-toggleBelow"
          />
        ) : null}
      </div>

      {hasChildren && isExpanded ? (
        <ul
          className={
            useVerticalLeafBranch
              ? 'sp-orgchart-children sp-orgchart-childrenVertical'
              : 'sp-orgchart-children'
          }
          role="group"
        >
          {node.children.map((child: IOrgChartNode) => (
            <OrgChartNodeView
              key={child.id}
              node={child}
              level={level + 1}
              expandedKeys={expandedKeys}
              onToggle={onToggle}
              resolvedProps={resolvedProps}
            />
          ))}
        </ul>
      ) : null}
    </li>
  );
};

export const SPOrgChart: React.FC<ISPOrgChartProps> = (props) => {
  const resolvedProps: Required<ISPOrgChartProps> = {
    items: props.items || [],
    rootPersonId: props.rootPersonId ?? '',
    rootPosition: props.rootPosition ?? '',
    rootEmail: props.rootEmail ?? '',
    rootLoginName: props.rootLoginName ?? '',
    idField: props.idField || DEFAULT_PROPS.idField,
    parentIdField: props.parentIdField || DEFAULT_PROPS.parentIdField,
    nameField: props.nameField || DEFAULT_PROPS.nameField,
    positionField: props.positionField || DEFAULT_PROPS.positionField,
    photoField: props.photoField || DEFAULT_PROPS.photoField,
    emailField: props.emailField || DEFAULT_PROPS.emailField,
    loginNameField: props.loginNameField || DEFAULT_PROPS.loginNameField,
    linkField: props.linkField || DEFAULT_PROPS.linkField,
    initiallyExpanded: props.initiallyExpanded ?? DEFAULT_PROPS.initiallyExpanded,
    expandAll: props.expandAll ?? DEFAULT_PROPS.expandAll,
    indentSize: props.indentSize ?? DEFAULT_PROPS.indentSize,
    emptyMessage: props.emptyMessage || DEFAULT_PROPS.emptyMessage
  };

  const { nodes, roots } = React.useMemo(
    () => {
      const preparedItems = materializeItems(resolvedProps.items, [
        resolvedProps.idField,
        resolvedProps.parentIdField,
        resolvedProps.nameField,
        resolvedProps.positionField,
        resolvedProps.photoField,
        resolvedProps.emailField,
        resolvedProps.loginNameField,
        resolvedProps.linkField,
        'Manager',
        'Department',
        'OfficeLocation',
        'Availability'
      ]);

      return buildTree(preparedItems, resolvedProps.idField, resolvedProps.parentIdField);
    },
    [
      resolvedProps.items,
      resolvedProps.idField,
      resolvedProps.parentIdField,
      resolvedProps.nameField,
      resolvedProps.positionField,
      resolvedProps.photoField,
      resolvedProps.emailField,
      resolvedProps.loginNameField,
      resolvedProps.linkField
    ]
  );

  const rootNode = React.useMemo(
    () => findRootNode(nodes, roots, resolvedProps),
    [nodes, roots, resolvedProps]
  );

  const [expandedKeys, setExpandedKeys] = React.useState<Record<string, boolean>>({});

  React.useEffect(() => {
    if (!rootNode) {
      setExpandedKeys({});
      return;
    }

    if (!resolvedProps.initiallyExpanded) {
      setExpandedKeys({});
      return;
    }

    const initialKeys: Record<string, boolean> = {};

    if (resolvedProps.expandAll) {
      collectExpandableNodes(rootNode, initialKeys);
    } else if (rootNode.children.length > 0) {
      initialKeys[rootNode.id] = true;
    }

    setExpandedKeys(initialKeys);
  }, [rootNode, resolvedProps.initiallyExpanded, resolvedProps.expandAll]);

  const handleToggle = React.useCallback((nodeId: string) => {
    setExpandedKeys((previous: Record<string, boolean>) => ({
      ...previous,
      [nodeId]: previous[nodeId] === false ? true : false
    }));
  }, []);

  if (!resolvedProps.items.length) {
    return <Text>{resolvedProps.emptyMessage}</Text>;
  }

  if (!rootNode) {
    return <Text>Unable to find the requested root person or position.</Text>;
  }

  return (
    <div className="sp-orgchart-tree" role="tree" aria-label="Organization chart">
      <ul className="sp-orgchart-rootList">
        <OrgChartNodeView
          node={rootNode}
          level={0}
          expandedKeys={expandedKeys}
          onToggle={handleToggle}
          resolvedProps={resolvedProps}
        />
      </ul>
    </div>
  );
};
