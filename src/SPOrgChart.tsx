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
  data: IOrgChartNodeData;
}

interface IOrgChartNodeData {
  name: string;
  position: string;
  email?: string;
  photoUrl: string;
  linkUrl?: string;
  managerName: string;
  department: string;
  officeLocation: string;
  presenceMeta: { label: string; className: string };
  teamsChatUrl: string;
}

export interface ISPOrgChartProps {
  items?: any[];
  rootPersonId?: string | number;
  rootPersonIds?: Array<string | number>;
  rootPosition?: string;
  rootPositions?: string[];
  rootEmail?: string;
  rootEmails?: string[];
  rootLoginName?: string;
  rootLoginNames?: string[];
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
  expandAllThreshold?: number;
  indentSize?: number;
  emptyMessage?: string;
  valueTransforms?: Record<
    string,
    {
      pattern: string;
      replacement: string;
      flags?: string;
    }
  >;
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
  expandAllThreshold: 150,
  indentSize: 24,
  emptyMessage: 'No organization data available.',
  valueTransforms: {}
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

const getCandidateFieldValues = (item: any, fieldNames: string[]): string[] => {
  const values = new Set<string>();

  fieldNames.forEach((fieldName: string) => {
    const value = normalizeValue(getFieldValue(item, fieldName));

    if (value) {
      values.add(value.toLowerCase());
    }
  });

  return Array.from(values);
};

const applyFieldTransform = (
  value: string | undefined,
  fieldName: string | undefined,
  transforms: Required<ISPOrgChartProps>['valueTransforms']
): string | undefined => {
  if (!value || !fieldName) {
    return value;
  }

  const transform = transforms[fieldName];

  if (!transform?.pattern) {
    return value;
  }

  try {
    return value.replace(new RegExp(transform.pattern, transform.flags || ''), transform.replacement);
  } catch (_error) {
    return value;
  }
};

const getResolvedFieldText = (
  item: any,
  fieldName: string | undefined,
  transforms: Required<ISPOrgChartProps>['valueTransforms']
): string | undefined => {
  const value = normalizeValue(getFieldValue(item, fieldName));

  return applyFieldTransform(value, fieldName, transforms);
};

const buildTeamsChatUrl = (email?: string): string => {
  if (!email) {
    return '';
  }

  return `https://teams.microsoft.com/l/chat/0/0?users=${encodeURIComponent(email)}`;
};

const buildProfilePhotoUrl = (email?: string): string => {
  if (!email) {
    return '';
  }

  return `/_layouts/15/userphoto.aspx?size=L&accountname=${encodeURIComponent(email)}`;
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

const resolveNodeData = (
  item: any,
  props: Required<ISPOrgChartProps>
): IOrgChartNodeData => {
  const name = getResolvedFieldText(item, props.nameField, props.valueTransforms) || 'Unnamed person';
  const position = getResolvedFieldText(item, props.positionField, props.valueTransforms) || '';
  const email = getResolvedFieldText(item, props.emailField, props.valueTransforms);
  const photoUrl =
    getResolvedFieldText(item, props.photoField, props.valueTransforms) ||
    buildProfilePhotoUrl(email || undefined);
  const linkUrl = getResolvedFieldText(item, props.linkField, props.valueTransforms);
  const manager = getFieldValue(item, 'Manager');
  const managerName =
    getResolvedFieldText(manager, props.nameField, props.valueTransforms) ||
    getResolvedFieldText(manager, 'DisplayName', props.valueTransforms) ||
    '';
  const department = getResolvedFieldText(item, 'Department', props.valueTransforms) || '';
  const officeLocation = getResolvedFieldText(item, 'OfficeLocation', props.valueTransforms) || '';
  const presenceMeta = getPresenceMeta(getFieldValue(item, 'Availability'));

  return {
    name,
    position,
    email: email || undefined,
    photoUrl,
    linkUrl: linkUrl || undefined,
    managerName,
    department,
    officeLocation,
    presenceMeta,
    teamsChatUrl: buildTeamsChatUrl(email || undefined)
  };
};

const buildTree = (
  items: any[],
  idField: string,
  parentIdField: string,
  props: Required<ISPOrgChartProps>
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
      children: [],
      data: resolveNodeData(item, props)
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

const findMatchingNodes = (
  nodes: Map<string, IOrgChartNode>,
  matcher: (node: IOrgChartNode) => boolean
): IOrgChartNode[] => {
  const matches: IOrgChartNode[] = [];
  const seenNodeIds = new Set<string>();

  Array.from(nodes.values()).forEach((node: IOrgChartNode) => {
    if (!seenNodeIds.has(node.id) && matcher(node)) {
      seenNodeIds.add(node.id);
      matches.push(node);
    }
  });

  return matches;
};

const findRootNodes = (
  nodes: Map<string, IOrgChartNode>,
  roots: IOrgChartNode[],
  props: Required<ISPOrgChartProps>
): IOrgChartNode[] => {
  const rootPersonIds = props.rootPersonIds.length
    ? props.rootPersonIds
    : props.rootPersonId !== undefined && props.rootPersonId !== null && props.rootPersonId !== ''
      ? [props.rootPersonId]
      : [];

  if (rootPersonIds.length) {
    return rootPersonIds
      .map((rootPersonId: string | number) => nodes.get(String(rootPersonId)))
      .filter((node: IOrgChartNode | undefined): node is IOrgChartNode => Boolean(node));
  }

  const rootEmails = props.rootEmails.length ? props.rootEmails : props.rootEmail ? [props.rootEmail] : [];

  if (rootEmails.length) {
    const normalizedRootEmails = rootEmails.map((rootEmail: string) => rootEmail.toLowerCase());

    return findMatchingNodes(nodes, (node: IOrgChartNode) => {
      const candidateEmails = getCandidateFieldValues(node.item, [
        props.emailField,
        'Email',
        'EMail',
        'email',
        'User/Email',
        'User/EMail',
        'User/email',
        'User/UserPrincipalName',
        'UserPrincipalName',
        'LoginName',
        'User/LoginName'
      ]);

      return candidateEmails.some((candidateEmail: string) => normalizedRootEmails.includes(candidateEmail));
    });
  }

  const rootLoginNames = props.rootLoginNames.length
    ? props.rootLoginNames
    : props.rootLoginName
      ? [props.rootLoginName]
      : [];

  if (rootLoginNames.length) {
    const normalizedRootLoginNames = rootLoginNames.map((rootLoginName: string) => rootLoginName.toLowerCase());

    return findMatchingNodes(nodes, (node: IOrgChartNode) => {
      const value = normalizeValue(getFieldValue(node.item, props.loginNameField));

      return value ? normalizedRootLoginNames.includes(value.toLowerCase()) : false;
    });
  }

  const rootPositions = props.rootPositions.length ? props.rootPositions : props.rootPosition ? [props.rootPosition] : [];

  if (rootPositions.length) {
    const normalizedRootPositions = rootPositions.map((rootPosition: string) => rootPosition.toLowerCase());

    return findMatchingNodes(nodes, (node: IOrgChartNode) => {
      const value = getResolvedFieldText(node.item, props.positionField, props.valueTransforms);

      return value ? normalizedRootPositions.includes(value.toLowerCase()) : false;
    });
  }

  return roots.length ? [roots[0]] : [];
};

const collectExpandableNodes = (node: IOrgChartNode, result: Record<string, boolean>): void => {
  if (node.children.length > 0) {
    result[node.id] = true;
  }

  node.children.forEach((child: IOrgChartNode) => collectExpandableNodes(child, result));
};

const countNodes = (node: IOrgChartNode): number => {
  let total = 1;

  node.children.forEach((child: IOrgChartNode) => {
    total += countNodes(child);
  });

  return total;
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

const zoomButtonStyles = {
  root: {
    width: 36,
    height: 36,
    borderRadius: 10,
    border: '1px solid #c8c8c8',
    backgroundColor: '#ffffff'
  }
};

const MIN_ZOOM = 0.6;
const MAX_ZOOM = 1.6;
const ZOOM_STEP = 0.1;

const OrgChartNodeView: React.FC<{
  node: IOrgChartNode;
  level: number;
  expandedKeys: Record<string, boolean>;
  onToggle: (nodeId: string) => void;
  onOpenDetails: (node: IOrgChartNode) => void;
  resolvedProps: Required<ISPOrgChartProps>;
}> = React.memo(({ node, level, expandedKeys, onToggle, onOpenDetails, resolvedProps }) => {
  const hasChildren = node.children.length > 0;
  const isExpanded = expandedKeys[node.id] !== false;
  const useVerticalLeafBranch =
    hasChildren && node.children.every((child: IOrgChartNode) => child.children.length === 0);
  const { name, position, email, photoUrl, linkUrl, presenceMeta, teamsChatUrl } = node.data;

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
              onOpenDetails(node);
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
              onOpenDetails={onOpenDetails}
              resolvedProps={resolvedProps}
            />
          ))}
        </ul>
      ) : null}
    </li>
  );
}, (previousProps, nextProps) => {
  return (
    previousProps.node === nextProps.node &&
    previousProps.level === nextProps.level &&
    previousProps.expandedKeys[previousProps.node.id] === nextProps.expandedKeys[nextProps.node.id] &&
    previousProps.onToggle === nextProps.onToggle &&
    previousProps.onOpenDetails === nextProps.onOpenDetails &&
    previousProps.resolvedProps === nextProps.resolvedProps
  );
});

export const SPOrgChart: React.FC<ISPOrgChartProps> = (props) => {
  const containerRef = React.useRef<HTMLDivElement | null>(null);
  const dragStateRef = React.useRef<{
    pointerId: number;
    startX: number;
    startY: number;
    scrollLeft: number;
    scrollTop: number;
    moved: boolean;
  } | null>(null);
  const suppressClickRef = React.useRef<boolean>(false);
  const resolvedProps: Required<ISPOrgChartProps> = React.useMemo(() => ({
    items: props.items || [],
    rootPersonId: props.rootPersonId ?? '',
    rootPersonIds: props.rootPersonIds ?? [],
    rootPosition: props.rootPosition ?? '',
    rootPositions: props.rootPositions ?? [],
    rootEmail: props.rootEmail ?? '',
    rootEmails: props.rootEmails ?? [],
    rootLoginName: props.rootLoginName ?? '',
    rootLoginNames: props.rootLoginNames ?? [],
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
    expandAllThreshold: props.expandAllThreshold ?? DEFAULT_PROPS.expandAllThreshold,
    indentSize: props.indentSize ?? DEFAULT_PROPS.indentSize,
    emptyMessage: props.emptyMessage || DEFAULT_PROPS.emptyMessage,
    valueTransforms: props.valueTransforms || DEFAULT_PROPS.valueTransforms
  }), [
    props.items,
    props.rootPersonId,
    props.rootPersonIds,
    props.rootPosition,
    props.rootPositions,
    props.rootEmail,
    props.rootEmails,
    props.rootLoginName,
    props.rootLoginNames,
    props.idField,
    props.parentIdField,
    props.nameField,
    props.positionField,
    props.photoField,
    props.emailField,
    props.loginNameField,
    props.linkField,
    props.initiallyExpanded,
    props.expandAll,
    props.expandAllThreshold,
    props.indentSize,
    props.emptyMessage,
    props.valueTransforms
  ]);

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

      return buildTree(preparedItems, resolvedProps.idField, resolvedProps.parentIdField, resolvedProps);
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
      resolvedProps.linkField,
      resolvedProps.valueTransforms
    ]
  );

  const rootNodes = React.useMemo(
    () => findRootNodes(nodes, roots, resolvedProps),
    [nodes, roots, resolvedProps]
  );

  const [expandedKeys, setExpandedKeys] = React.useState<Record<string, boolean>>({});
  const [selectedNode, setSelectedNode] = React.useState<IOrgChartNode | undefined>();
  const [zoom, setZoom] = React.useState<number>(1);

  const shouldExpandAllInitially = React.useMemo(() => {
    if (props.expandAll === true) {
      return true;
    }

    if (!resolvedProps.expandAll) {
      return false;
    }

    let totalNodes = 0;

    rootNodes.forEach((rootNode: IOrgChartNode) => {
      totalNodes += countNodes(rootNode);
    });

    return totalNodes <= resolvedProps.expandAllThreshold;
  }, [rootNodes, resolvedProps.expandAll, resolvedProps.expandAllThreshold]);

  React.useEffect(() => {
    if (!rootNodes.length) {
      setExpandedKeys({});
      return;
    }

    if (!resolvedProps.initiallyExpanded) {
      setExpandedKeys({});
      return;
    }

    const initialKeys: Record<string, boolean> = {};

    if (shouldExpandAllInitially) {
      rootNodes.forEach((rootNode: IOrgChartNode) => collectExpandableNodes(rootNode, initialKeys));
    } else {
      rootNodes.forEach((rootNode: IOrgChartNode) => {
        if (rootNode.children.length > 0) {
          initialKeys[rootNode.id] = true;
        }
      });
    }

    setExpandedKeys(initialKeys);
  }, [rootNodes, resolvedProps.initiallyExpanded, shouldExpandAllInitially]);

  const handleToggle = React.useCallback((nodeId: string) => {
    setExpandedKeys((previous: Record<string, boolean>) => ({
      ...previous,
      [nodeId]: previous[nodeId] === false ? true : false
    }));
  }, []);

  const handleOpenDetails = React.useCallback((node: IOrgChartNode) => {
    setSelectedNode(node);
  }, []);

  const handleZoomIn = React.useCallback(() => {
    setZoom((previous: number) => Math.min(MAX_ZOOM, Number((previous + ZOOM_STEP).toFixed(2))));
  }, []);

  const handleZoomOut = React.useCallback(() => {
    setZoom((previous: number) => Math.max(MIN_ZOOM, Number((previous - ZOOM_STEP).toFixed(2))));
  }, []);

  const handleZoomReset = React.useCallback(() => {
    setZoom(1);
  }, []);

  const stopDragging = React.useCallback(() => {
    const container = containerRef.current;

    if (container) {
      container.classList.remove('sp-orgchart-treeDragging');
    }

    if (dragStateRef.current?.moved) {
      suppressClickRef.current = true;
    }

    dragStateRef.current = null;
  }, []);

  const handlePointerDown = React.useCallback((event: React.PointerEvent<HTMLDivElement>) => {
    const container = containerRef.current;
    const target = event.target as HTMLElement | null;

    if (!container || !target) {
      return;
    }

    if (target.closest('button, input, textarea, select, option, [role="button"]')) {
      return;
    }

    dragStateRef.current = {
      pointerId: event.pointerId,
      startX: event.clientX,
      startY: event.clientY,
      scrollLeft: container.scrollLeft,
      scrollTop: container.scrollTop,
      moved: false
    };

    container.classList.add('sp-orgchart-treeDragging');
    container.setPointerCapture(event.pointerId);
  }, []);

  const handlePointerMove = React.useCallback((event: React.PointerEvent<HTMLDivElement>) => {
    const container = containerRef.current;
    const dragState = dragStateRef.current;

    if (!container || !dragState || dragState.pointerId !== event.pointerId) {
      return;
    }

    if (!dragState.moved) {
      const movementX = Math.abs(event.clientX - dragState.startX);
      const movementY = Math.abs(event.clientY - dragState.startY);

      if (movementX > 4 || movementY > 4) {
        dragState.moved = true;
      }
    }

    container.scrollLeft = dragState.scrollLeft - (event.clientX - dragState.startX);
    container.scrollTop = dragState.scrollTop - (event.clientY - dragState.startY);
  }, []);

  const handlePointerUp = React.useCallback((event: React.PointerEvent<HTMLDivElement>) => {
    const container = containerRef.current;
    const dragState = dragStateRef.current;

    if (!container || !dragState || dragState.pointerId !== event.pointerId) {
      return;
    }

    container.releasePointerCapture(event.pointerId);
    stopDragging();
  }, [stopDragging]);

  const handlePointerCancel = React.useCallback((event: React.PointerEvent<HTMLDivElement>) => {
    const container = containerRef.current;
    const dragState = dragStateRef.current;

    if (container && dragState && dragState.pointerId === event.pointerId) {
      container.releasePointerCapture(event.pointerId);
    }

    stopDragging();
  }, [stopDragging]);

  const handleClickCapture = React.useCallback((event: React.MouseEvent<HTMLDivElement>) => {
    if (!suppressClickRef.current) {
      return;
    }

    suppressClickRef.current = false;
    event.preventDefault();
    event.stopPropagation();
  }, []);

  if (!resolvedProps.items.length) {
    return <Text>{resolvedProps.emptyMessage}</Text>;
  }

  if (!rootNodes.length) {
    return <Text>Unable to find the requested root person or position.</Text>;
  }

  return (
    <div
      ref={containerRef}
      className="sp-orgchart-tree"
      role="tree"
      aria-label="Organization chart"
      onPointerDown={handlePointerDown}
      onPointerMove={handlePointerMove}
      onPointerUp={handlePointerUp}
      onPointerCancel={handlePointerCancel}
      onPointerLeave={handlePointerCancel}
      onClickCapture={handleClickCapture}
    >
      <div className="sp-orgchart-controls" onPointerDown={(event) => event.stopPropagation()}>
        <IconButton
          iconProps={{ iconName: 'ZoomOut' }}
          title="Zoom out"
          ariaLabel="Zoom out"
          styles={zoomButtonStyles}
          disabled={zoom <= MIN_ZOOM}
          onClick={handleZoomOut}
        />
        <DefaultButton
          text={`${Math.round(zoom * 100)}%`}
          className="sp-orgchart-zoomLabel"
          onClick={handleZoomReset}
        />
        <IconButton
          iconProps={{ iconName: 'ZoomIn' }}
          title="Zoom in"
          ariaLabel="Zoom in"
          styles={zoomButtonStyles}
          disabled={zoom >= MAX_ZOOM}
          onClick={handleZoomIn}
        />
      </div>

      <div className="sp-orgchart-canvas" style={{ transform: `scale(${zoom})` }}>
        <ul className="sp-orgchart-rootList">
          {rootNodes.map((rootNode: IOrgChartNode) => (
            <OrgChartNodeView
              key={rootNode.id}
              node={rootNode}
              level={0}
              expandedKeys={expandedKeys}
              onToggle={handleToggle}
              onOpenDetails={handleOpenDetails}
              resolvedProps={resolvedProps}
            />
          ))}
        </ul>
      </div>

      <Dialog
        hidden={!selectedNode}
        onDismiss={() => setSelectedNode(undefined)}
        dialogContentProps={{
          title: ''
        }}
      >
        {selectedNode ? (
          <Stack tokens={{ childrenGap: 14 }}>
            <Persona
              text={selectedNode.data.name}
              secondaryText={selectedNode.data.position}
              imageUrl={selectedNode.data.photoUrl}
              size={PersonaSize.size72}
            />
            {selectedNode.data.email ? (
              <Stack tokens={{ childrenGap: 4 }}>
                <Text variant="mediumPlus">Email</Text>
                <Link href={`mailto:${selectedNode.data.email}`}>{selectedNode.data.email}</Link>
              </Stack>
            ) : null}
            {selectedNode.data.managerName ? (
              <Stack tokens={{ childrenGap: 4 }}>
                <Text variant="mediumPlus">Manager</Text>
                <Text>{selectedNode.data.managerName}</Text>
              </Stack>
            ) : null}
            {selectedNode.data.department ? (
              <Stack tokens={{ childrenGap: 4 }}>
                <Text variant="mediumPlus">Department</Text>
                <Text>{selectedNode.data.department}</Text>
              </Stack>
            ) : null}
            {selectedNode.data.officeLocation ? (
              <Stack tokens={{ childrenGap: 4 }}>
                <Text variant="mediumPlus">Office</Text>
                <Text>{selectedNode.data.officeLocation}</Text>
              </Stack>
            ) : null}
          </Stack>
        ) : null}
      </Dialog>
    </div>
  );
};
