import * as React from 'react';
import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/site-users";
import styles from './Orgchart.module.scss';
import './style.css';

interface IOrgchartProps {
  description: string;
  siteurl: string;
  context: any;
  employeeCount: number;
}

interface IEmployee {
  Id: number;
  Title: string;
  EmployeeName: { Title: string; EMail: string; PictureUrl: string } | null;
  EmployeeManagers: { Title: string; EMail: string }[] | null;
  Department: string | null;
  Phone: number;
  Location: string;
  Views?: { Title: string }[] | null;
  ViewName?: string[] | string;
  children?: IEmployee[];
  level?: number;
  displayOrder?: number;
  isNoDepartment?: boolean;
  isImmediateChildOfTopLevel?: boolean;
  Color?: string | null;
  FontColor?: string | null;
  ParentOrder?: number;
  ChildOrder?: number;
  SubChildOrder?: number;
}

interface IOrgView {
  Id: number;
  Title: string;
}

interface IOrgchartState {
  employees: IEmployee[];
  expandedNodes: Set<number>;
  loading: boolean;
  error: string | null;
  currentemployeeCount: number;
  allEmployees: IEmployee[];
  managerPlaceholders: IEmployee[];
  flattenedEmployees: IEmployee[];
  searchTerm: string;
  selectedDirector: number | null;
  selectedCard: number | null;
  originalEmployees: IEmployee[];
  currentScreen: 'Screen0' | 'Screen1' | 'Screen2';
  orgViews: IOrgView[];
  selectedView: string | null;
  selectedOrgView: string | null;
  showPopup: boolean;
  selectedEmployee: IEmployee | null;
  popupPosition: { top: number; left: number } | null;
  zoomLevel: number;
}

export default class Orgchart extends React.Component<IOrgchartProps, IOrgchartState> {
  private orgchartWrapperRef: React.RefObject<HTMLDivElement>;

  constructor(props: IOrgchartProps) {
    super(props);
    this.orgchartWrapperRef = React.createRef();
    this.state = {
      employees: [],
      expandedNodes: new Set(),
      loading: true,
      error: null,
      currentemployeeCount: props.employeeCount,
      allEmployees: [],
      managerPlaceholders: [],
      flattenedEmployees: [],
      searchTerm: '',
      selectedDirector: null,
      selectedCard: null,
      originalEmployees: [],
      currentScreen: 'Screen0',
      orgViews: [],
      selectedView: null,
      selectedOrgView: null,
      showPopup: false,
      selectedEmployee: null,
      popupPosition: null,
      zoomLevel: 1,
    };
  }

  public async componentDidMount(): Promise<void> {
    await this.fetchOrgViews();
    await this.fetchEmployees();
  }

  public componentDidUpdate(prevProps: IOrgchartProps, prevState: IOrgchartState): void {
    if (
      prevState.currentScreen !== this.state.currentScreen ||
      prevState.selectedView !== this.state.selectedView ||
      prevState.employees !== this.state.employees
    ) {
      this.setState({
        showPopup: false,
        selectedEmployee: null,
        selectedCard: null,
        popupPosition: null,
      });

      this.centerScrollPosition();
    }

    if (prevProps.employeeCount !== this.props.employeeCount) {
      this.setState({ currentemployeeCount: this.props.employeeCount }, () => {
        this.filterEmployeesBasedOnCount();
      });
    }
  }

  private centerScrollPosition(): void {
    const wrapper = this.orgchartWrapperRef.current;
    if (wrapper) {
      const scrollWidth = wrapper.scrollWidth;
      const clientWidth = wrapper.clientWidth;
      const centerPosition = (scrollWidth - clientWidth) / 2;
      wrapper.scrollLeft = centerPosition;
    }
  }

  public async fetchOrgViews(): Promise<void> {
    try {
      const items = await sp.web.lists.getByTitle("OrgViews").items
        .select('Id', 'Title')
        .getAll();
      console.log('Fetched OrgViews:', items);
      this.setState({
        orgViews: items,
      });
    } catch (error) {
      console.error('Error fetching OrgViews:', error);
      this.setState({ error: 'Failed to fetch views', loading: false });
    }
  }

  public assignDisplayOrder(employees: IEmployee[]): IEmployee[] {
    let order = 0;

    const assignOrder = (employee: IEmployee, level: number = 0): void => {
      employee.level = level;
      employee.displayOrder = order++;

      if (employee.children) {
        employee.children.sort((a, b) => {
          const parentOrderA = a.ParentOrder || 0;
          const parentOrderB = b.ParentOrder || 0;
          const childOrderA = a.ChildOrder || 0;
          const childOrderB = b.ChildOrder || 0;
          const subChildOrderA = a.SubChildOrder || 0;
          const subChildOrderB = b.SubChildOrder || 0;
          return parentOrderA - parentOrderB || childOrderA - childOrderB || subChildOrderA - subChildOrderB || a.Id - b.Id;
        });
        employee.children.forEach(child => assignOrder(child, level + 1));
      }
    };

    const highAuthorityRoles = ['Chief Executive Officer', 'Chief Engineering Officer', 'Chief Operating Officer'];
    const sortedEmployees = employees.sort((a, b) => {
      const isHighAuthorityA = a.EmployeeManagers === null && highAuthorityRoles.some(role => a.Title?.toLowerCase().includes(role.toLowerCase()));
      const isHighAuthorityB = b.EmployeeManagers === null && highAuthorityRoles.some(role => b.Title?.toLowerCase().includes(role.toLowerCase()));
      if (isHighAuthorityA && isHighAuthorityB) {
        return (a.ParentOrder || 0) - (b.ParentOrder || 0) || a.Id - b.Id;
      } else if (isHighAuthorityA) {
        return -1;
      } else if (isHighAuthorityB) {
        return 1;
      }
      return (a.ParentOrder || 0) - (b.ParentOrder || 0) || a.Id - b.Id;
    });

    sortedEmployees.forEach(emp => assignOrder(emp));
    return sortedEmployees;
  }

  public flattenHierarchy(employees: IEmployee[]): IEmployee[] {
    let flattened: IEmployee[] = [];
    const seen = new Set<number>();

    const flatten = (employee: IEmployee, level: number = 0): void => {
      if (seen.has(employee.Id)) return;
      seen.add(employee.Id);

      const employeeWithLevel = { ...employee, level };
      flattened.push(employeeWithLevel);

      if (employee.children && employee.children.length > 0) {
        employee.children.forEach(child => flatten(child, level + 1));
      }
    };

    employees.forEach(emp => flatten(emp));
    console.log('Flattened Hierarchy:', flattened.map(emp => ({
      Id: emp.Id,
      Name: emp.EmployeeName?.Title,
      Children: emp.children?.map(child => child.Id) || [],
      isImmediateChildOfTopLevel: emp.isImmediateChildOfTopLevel
    })));
    return flattened;
  }

  public async fetchEmployees(): Promise<void> {
    try {
      const items = await sp.web.lists.getByTitle("OrgChart").items
        .select(
          'Id',
          'Title',
          'EmployeeName/Title',
          'EmployeeName/EMail',
          'EmployeeManagers/Id',
          'Department',
          'Phone',
          'Location',
          'Views/Title',
          'ViewName',
          'Color',
          'FontColor',
          'ParentOrder',
          'ChildOrder',
          'SubChildOrder'
        )
        .expand('EmployeeName', 'EmployeeManagers', 'Views')
        .getAll();

      console.log('Fetched OrgChart Items:', items);

      const managerIds = new Set<number>();
      items.forEach(item => {
        if (item.EmployeeManagers && item.EmployeeManagers.length > 0) {
          item.EmployeeManagers.forEach((mgr: any) => {
            if (mgr.Id) {
              managerIds.add(mgr.Id);
            }
          });
        }
      });

      let managersDetails: { [key: number]: { Title: string; EMail: string } } = {};
      if (managerIds.size > 0) {
        const managerIdsArray = Array.from(managerIds);
        const managerDetails = await sp.web.siteUsers
          .select('Id', 'Title', 'Email')
          .filter(managerIdsArray.map(id => `Id eq ${id}`).join(' or '))
          .get();

        managerDetails.forEach(user => {
          managersDetails[user.Id] = {
            Title: user.Title || 'N/A',
            EMail: user.Email || 'N/A'
          };
        });
      }

      console.log('Fetched Manager Details:', managersDetails);

      const processedItems = items.map(item => {
        let managers: { Title: string; EMail: string }[] | null = null;

        if (item.EmployeeManagers && item.EmployeeManagers.length > 0) {
          managers = item.EmployeeManagers.map((mgr: any) => {
            const mgrDetails = managersDetails[mgr.Id] || { Title: 'Unknown', EMail: 'N/A' };
            return {
              Title: mgrDetails.Title,
              EMail: mgrDetails.EMail
            };
          });
        }

        return {
          Id: item.Id,
          Title: item.Title || 'N/A',
          EmployeeName: item.EmployeeName ? {
            Title: item.EmployeeName.Title || 'N/A',
            EMail: item.EmployeeName.EMail || 'N/A',
            PictureUrl: `${this.props.siteurl}/_layouts/15/userphoto.aspx?size=L&accountname=${item.EmployeeName?.EMail || ''}`
          } : null,
          EmployeeManagers: managers,
          Department: item.Department || null,
          Phone: item.Phone || 0,
          Location: item.Location || '',
          Views: item.Views,
          ViewName: item.ViewName,
          Color: item.Color || '#dadbdc',
          FontColor: item.FontColor || '#000000',
          ParentOrder: item.ParentOrder || 0,
          ChildOrder: item.ChildOrder || 0,
          SubChildOrder: item.SubChildOrder || 0,
          isImmediateChildOfTopLevel: false,
        };
      });

      console.log('Processed OrgChart Items:', processedItems);

      const regularEmployees = processedItems.filter(item => item.EmployeeName);
      const managerPlaceholders = processedItems.filter(item => !item.EmployeeName);

      console.log('Regular Employees:', regularEmployees);
      console.log('Manager Placeholders:', managerPlaceholders);

      const hierarchy = this.buildHierarchy(regularEmployees);
      console.log('Built Hierarchy (Executives):', hierarchy);

      hierarchy.forEach(exec => {
        console.log(`Executive: ${exec.EmployeeName?.Title} (ID: ${exec.Id})`);
        console.log('Children:', exec.children?.map(child => ({
          Id: child.Id,
          Name: child.EmployeeName?.Title,
          Title: child.Title,
          isImmediateChildOfTopLevel: child.isImmediateChildOfTopLevel
        })));
      });

      const orderedHierarchy = this.assignDisplayOrder(hierarchy);
      const flattened = this.flattenHierarchy(orderedHierarchy);
      const allFlattened = [...flattened, ...managerPlaceholders];

      this.setState({
        allEmployees: orderedHierarchy,
        managerPlaceholders: managerPlaceholders,
        flattenedEmployees: allFlattened,
        originalEmployees: [...orderedHierarchy],
        loading: false
      }, () => {
        this.filterEmployeesBasedOnCount();
        this.centerScrollPosition();
      });

    } catch (error) {
      console.error('Error fetching employees:', error);
      this.setState({ error: 'Failed to fetch data', loading: false });
    }
  }

  public buildHierarchy(items: IEmployee[]): IEmployee[] {
    const employeeMap = new Map<number, IEmployee>();
    const executives: IEmployee[] = [];
    const seenEmployees = new Set<number>();

    items.forEach(item => {
      if (!employeeMap.has(item.Id)) {
        const employee = { ...item, children: [], isImmediateChildOfTopLevel: false };
        employeeMap.set(item.Id, employee);
      }
    });

    items.forEach(item => {
      const employee = employeeMap.get(item.Id);
      if (employee && !seenEmployees.has(employee.Id)) {
        if (!employee.EmployeeManagers || employee.EmployeeManagers.length === 0) {
          executives.push(employee);
          seenEmployees.add(employee.Id);
          if (employee.children) {
            employee.children.forEach(child => {
              const childEmployee = employeeMap.get(child.Id);
              if (childEmployee) {
                childEmployee.isImmediateChildOfTopLevel = true;
              }
            });
          }
        }
      }
    });

    items.forEach(item => {
      const employee = employeeMap.get(item.Id);
      if (employee && !seenEmployees.has(employee.Id)) {
        let addedToManager = false;
        if (employee.EmployeeManagers && employee.EmployeeManagers.length > 0) {
          const primaryManager = employee.EmployeeManagers[0];
          const primaryManagerEmp = Array.from(employeeMap.values()).find(
            emp => emp.EmployeeName?.Title === primaryManager.Title
          );

          if (primaryManagerEmp) {
            if (!primaryManagerEmp.children) {
              primaryManagerEmp.children = [];
            }
            if (!primaryManagerEmp.children.some(child => child.Id === employee.Id)) {
              primaryManagerEmp.children.push(employee);
              seenEmployees.add(employee.Id);
              addedToManager = true;
              if (!primaryManagerEmp.EmployeeManagers || primaryManagerEmp.EmployeeManagers.length === 0) {
                employee.isImmediateChildOfTopLevel = true;
              }
            }
          }
        }

        if (!addedToManager && !seenEmployees.has(employee.Id)) {
          executives.push(employee);
          seenEmployees.add(employee.Id);
          if (employee.children) {
            employee.children.forEach(child => {
              const childEmployee = employeeMap.get(child.Id);
              if (childEmployee) {
                childEmployee.isImmediateChildOfTopLevel = true;
              }
            });
          }
        }
      }
    });

    const finalExecutives: IEmployee[] = [];
    const executiveIds = new Set<number>();

    executives.forEach(exec => {
      if (!executiveIds.has(exec.Id)) {
        finalExecutives.push(exec);
        executiveIds.add(exec.Id);
      }
    });

    const cleanHierarchy = (employee: IEmployee, seen: Set<number> = new Set()): IEmployee => {
      if (seen.has(employee.Id)) {
        return { ...employee, children: [] };
      }
      seen.add(employee.Id);

      if (employee.children) {
        const uniqueChildren: IEmployee[] = [];
        const childIds = new Set<number>();

        employee.children.forEach(child => {
          if (!childIds.has(child.Id)) {
            childIds.add(child.Id);
            uniqueChildren.push(cleanHierarchy(child, seen));
          }
        });

        employee.children = uniqueChildren;
      }

      return employee;
    };

    const cleanedHierarchy = finalExecutives.map(exec => cleanHierarchy(exec));
    console.log('Cleaned Hierarchy:', cleanedHierarchy.map(exec => ({
      Id: exec.Id,
      Name: exec.EmployeeName?.Title,
      Children: exec.children?.map(child => child.Id) || [],
      isImmediateChildOfTopLevel: exec.isImmediateChildOfTopLevel
    })));
    return cleanedHierarchy;
  }

  public handleSearchChange = (event: React.ChangeEvent<HTMLInputElement>): void => {
    const searchTerm = event.target.value;
    this.setState({ searchTerm }, () => {
      this.filterEmployeesBasedOnSearch();
    });
  };

  public findParent(employeeId: number, employees: IEmployee[]): IEmployee | undefined {
    for (const emp of employees) {
      if (emp.children) {
        const directParent = emp.children.find(child => child.Id === employeeId);
        if (directParent) return emp;
        const nestedParent = this.findParent(employeeId, emp.children);
        if (nestedParent) return nestedParent;
      }
    }
    return undefined;
  }

  public getAncestorIds(employee: IEmployee): number[] {
    const ancestorIds: number[] = [];
    let current: IEmployee | undefined = employee;

    while (current) {
      const parent = this.findParent(current.Id, this.state.allEmployees);
      if (parent) {
        ancestorIds.push(parent.Id);
        current = parent;
      } else {
        current = undefined;
      }
    }

    return ancestorIds;
  }

  public matchesView(emp: IEmployee, normalizedSelectedView: string): boolean {
    const matchesViewName = Array.isArray(emp.ViewName)
      ? emp.ViewName.some(name => name.toLowerCase().trim() === normalizedSelectedView)
      : (typeof emp.ViewName === 'string' && emp.ViewName.toLowerCase().trim() === normalizedSelectedView);
    const matchesViews = emp.Views
      ? emp.Views.some(view => {
          const viewTitle = view.Title?.toLowerCase().trim();
          return viewTitle ? (viewTitle === normalizedSelectedView || viewTitle.includes(normalizedSelectedView)) : false;
        })
      : false;
    return matchesViewName || matchesViews;
  }

  public filterHierarchyByView(employees: IEmployee[], selectedView: string): IEmployee[] {
    const normalizedSelectedView = selectedView.toLowerCase().trim();

    const filterEmployee = (emp: IEmployee): IEmployee | null => {
      const employeeMatches = this.matchesView(emp, normalizedSelectedView);

      const filteredChildren = emp.children
        ? emp.children
            .map(child => filterEmployee(child))
            .filter((child): child is IEmployee => child !== null)
        : undefined;

      if (employeeMatches || (filteredChildren && filteredChildren.length > 0)) {
        const result: IEmployee = { ...emp };
        if (filteredChildren && filteredChildren.length > 0) {
          result.children = filteredChildren;
        }
        return result;
      }

      return null;
    };

    const filteredHierarchy = employees
      .map(emp => filterEmployee(emp))
      .filter((emp): emp is IEmployee => emp !== null);
    return filteredHierarchy;
  }

  public filterEmployeesBasedOnSearch(): void {
    const { searchTerm, allEmployees, selectedView } = this.state;

    if (!searchTerm && !selectedView) {
      this.filterEmployeesBasedOnCount();
      return;
    }

    let filteredEmployees: IEmployee[] = [...allEmployees];

    if (selectedView) {
      filteredEmployees = this.filterHierarchyByView(filteredEmployees, selectedView);
    }

    if (searchTerm) {
      const lowerSearchTerm = searchTerm.toLowerCase().trim();

      const allFlattened = this.flattenHierarchy(filteredEmployees);
      const nameMatches = allFlattened.filter(emp =>
        emp.EmployeeName?.Title?.toLowerCase().includes(lowerSearchTerm)
      );

      if (nameMatches.length > 0) {
        filteredEmployees = nameMatches.map(emp => ({
          ...emp,
          children: [],
          EmployeeManagers: null,
          isImmediateChildOfTopLevel: false
        }));
      } else {
        const filterBySearch = (emp: IEmployee): IEmployee | null => {
          const matchesSearch =
            emp.Department?.toLowerCase().includes(lowerSearchTerm) ||
            emp.Location?.toLowerCase().includes(lowerSearchTerm);

          const filteredChildren = emp.children
            ? emp.children
                .map(child => filterBySearch(child))
                .filter((child): child is IEmployee => child !== null)
            : undefined;

          if (matchesSearch || (filteredChildren && filteredChildren.length > 0)) {
            const result: IEmployee = { ...emp };
            if (filteredChildren && filteredChildren.length > 0) {
              result.children = filteredChildren;
            }
            return result;
          }
          return null;
        };

        filteredEmployees = filteredEmployees
          .map(emp => filterBySearch(emp))
          .filter((emp): emp is IEmployee => emp !== null);
      }
    }

    this.setState({
      employees: filteredEmployees,
      expandedNodes: new Set(this.flattenHierarchy(filteredEmployees).map(emp => emp.Id))
    });
  }

  public filterEmployeesBasedOnCount(): void {
    const { currentemployeeCount, allEmployees, selectedView } = this.state;

    if (currentemployeeCount <= 0 || !allEmployees.length) {
      this.setState({ employees: [] });
      return;
    }

    let filteredEmployees: IEmployee[] = [...allEmployees];

    if (selectedView) {
      filteredEmployees = this.filterHierarchyByView(filteredEmployees, selectedView);
    }

    const flattened = this.flattenHierarchy(filteredEmployees).sort((a, b) =>
      (a.displayOrder || 0) - (b.displayOrder || 0)
    );
    const selectedIds = new Set(flattened.slice(0, currentemployeeCount).map(emp => emp.Id));

    const rebuildHierarchy = (employees: IEmployee[]): IEmployee[] => {
      return employees
        .map(emp => {
          if (!selectedIds.has(emp.Id)) {
            const filteredChildren = emp.children ? rebuildHierarchy(emp.children) : undefined;
            if (!filteredChildren || filteredChildren.length === 0) return null;
            const result: IEmployee = { ...emp };
            if (filteredChildren.length > 0) {
              result.children = filteredChildren;
            }
            return result;
          }
          const result: IEmployee = { ...emp };
          const rebuiltChildren = emp.children ? rebuildHierarchy(emp.children) : undefined;
          if (rebuiltChildren && rebuiltChildren.length > 0) {
            result.children = rebuiltChildren;
          }
          return result;
        })
        .filter((emp): emp is IEmployee => emp !== null);
    };

    filteredEmployees = rebuildHierarchy(filteredEmployees);
    this.setState({
      employees: filteredEmployees,
      expandedNodes: new Set(this.flattenHierarchy(filteredEmployees).map(emp => emp.Id))
    });
  }

  public rebuildFilteredHierarchy(selectedIds: number[], allEmployees: IEmployee[]): IEmployee[] {
    const seen = new Set<number>();
    const employeeMap = new Map<number, IEmployee>();
    const childIds = new Set<number>();

    const deduplicateEmployees = (employees: IEmployee[]): void => {
      employees.forEach(emp => {
        if (!employeeMap.has(emp.Id)) {
          employeeMap.set(emp.Id, { ...emp, children: [] });
        }
        if (emp.children) {
          emp.children.forEach(child => childIds.add(child.Id));
          deduplicateEmployees(emp.children);
        }
      });
    };

    deduplicateEmployees(allEmployees);

    const rebuild = (employees: IEmployee[], seenAtLevel: Set<number> = new Set()): IEmployee[] => {
      return employees
        .filter(emp => {
          if (!selectedIds.includes(emp.Id)) return false;
          if (seen.has(emp.Id) || seenAtLevel.has(emp.Id)) return false;
          if (!emp.EmployeeManagers || emp.EmployeeManagers.length === 0) {
            seen.add(emp.Id);
            seenAtLevel.add(emp.Id);
            return true;
          }
          if (childIds.has(emp.Id)) return false;
          seen.add(emp.Id);
          seenAtLevel.add(emp.Id);
          return true;
        })
        .map(emp => {
          const mappedEmp = employeeMap.get(emp.Id)!;
          return {
            ...mappedEmp,
            children: emp.children ? rebuild(emp.children, new Set()) : [],
            isImmediateChildOfTopLevel: mappedEmp.isImmediateChildOfTopLevel
          };
        });
    };

    const rebuiltHierarchy = rebuild(allEmployees);
    console.log('Rebuilt Hierarchy:', rebuiltHierarchy.map(emp => ({
      Id: emp.Id,
      Name: emp.EmployeeName?.Title,
      Children: emp.children?.map(child => child.Id) || [],
      isImmediateChildOfTopLevel: emp.isImmediateChildOfTopLevel
    })));
    return rebuiltHierarchy;
  }

  public toggleNode = (employeeId: number): void => {
    this.setState(prevState => {
      const newExpandedNodes = new Set(prevState.expandedNodes);
      if (newExpandedNodes.has(employeeId)) {
        newExpandedNodes.delete(employeeId);
      } else {
        newExpandedNodes.add(employeeId);
      }
      return { expandedNodes: newExpandedNodes };
    });
  }

  public getNodeClassName(employee: IEmployee): string {
    const baseClass = styles.node;
    const roleClass = this.getRoleClass(employee.Title);
    const expandedClass = this.state.expandedNodes.has(employee.Id) ? styles.expanded : '';
    const selectedClass = this.state.selectedCard === employee.Id ? styles.selectedCard : '';
    const topLevelClass = this.state.allEmployees.some(exec => exec.Id === employee.Id) ? styles.topLevel : '';
    const immediateChildClass = employee.isImmediateChildOfTopLevel ? styles.immediateChild : '';

    console.log(`Employee: ${employee.EmployeeName?.Title}, isImmediateChildOfTopLevel: ${employee.isImmediateChildOfTopLevel}, Class: ${immediateChildClass}`);

    return `${baseClass} ${roleClass} ${expandedClass} ${selectedClass} ${topLevelClass} ${immediateChildClass}`.trim();
  }

  public getRoleClass(title: string): string {
    const highAuthorityRoles = ['Chief Executive Officer', 'Chief Engineering Officer', 'Chief Operating Officer'];
    const safeTitle = title || '';
    return highAuthorityRoles.some(role => safeTitle.toLowerCase().includes(role))
      ? styles.executive
      : styles.default;
  }

  public findEmployeeById(employeeId: number, employees: IEmployee[] = this.state.allEmployees): IEmployee | null {
    for (const emp of employees) {
      if (emp.Id === employeeId) {
        return emp;
      }
      if (emp.children && emp.children.length > 0) {
        const foundInChildren = this.findEmployeeById(employeeId, emp.children);
        if (foundInChildren) {
          return foundInChildren;
        }
      }
    }
    return null;
  }

  public handleCardClick = (employee: IEmployee, event: React.MouseEvent) => {
    event.stopPropagation();

    const cardElement = event.currentTarget as HTMLElement;
    const rect = cardElement.getBoundingClientRect();
    const scrollY = window.scrollY || window.pageYOffset;
    const scrollX = window.scrollX || window.pageXOffset;
    const popupPosition = {
      top: rect.top + scrollY + rect.height / 2,
      left: rect.left + scrollX + rect.width / 2,
    };

    if (this.state.selectedCard === employee.Id) {
      this.setState({
        showPopup: !this.state.showPopup,
        selectedEmployee: employee,
        popupPosition,
      });
    } else {
      this.setState({
        showPopup: true,
        selectedEmployee: employee,
        selectedCard: employee.Id,
        popupPosition,
      });
    }
  };

  public closePopup = () => {
    this.setState({ showPopup: false, selectedEmployee: null, selectedCard: null, popupPosition: null });
  };

  public renderPopup = () => {
    const { showPopup, selectedEmployee, popupPosition } = this.state;

    if (!showPopup || !selectedEmployee || !popupPosition) return null;

    return (
      <div
        className={styles.popupContainer}
        style={{
          top: `${popupPosition.top}px`,
          left: `${popupPosition.left}px`,
        }}
      >
        <div className={styles.popupContent}>
          <div className={styles.popupHeader}>
            <span className={styles.closeIcon} onClick={this.closePopup}>×</span>
          </div>
          <div className={styles.popupBody}>
            <div className={styles.popupAvatarContainer}>
              <div
                className={styles.popupAvatar}
                style={{
                  backgroundImage: selectedEmployee.EmployeeName?.PictureUrl
                    ? `url('${selectedEmployee.EmployeeName.PictureUrl}')`
                    : `url('${this.props.siteurl}/SiteAssets/nouserimageicon.jpg')`,
                }}
              />
            </div>
            <div className={styles.popupInfo}>
              <div className={styles.popupName}>{selectedEmployee.EmployeeName?.Title || 'N/A'}</div>
              <div className={styles.popupTitle}>{selectedEmployee.Title || 'N/A'}</div>
              <div className={styles.popupEmail}>{selectedEmployee.EmployeeName?.EMail || 'N/A'}</div>
              <div className={styles.popupPhone}>{selectedEmployee.Phone || 'N/A'}</div>
              <div className={styles.popupLocation}>{selectedEmployee.Location || 'N/A'}</div>
            </div>
          </div>
        </div>
      </div>
    );
  };

  public renderChildrenHorizontally(children: IEmployee[], parentId: number): JSX.Element {
    if (!children || children.length === 0) return <></>;

    const uniqueChildren = Array.from(new Map(children.map(child => [child.Id, child])).values());

    console.log(`Rendering children for parent ${parentId}:`, uniqueChildren.map(child => ({
      Id: child.Id,
      Name: child.EmployeeName?.Title,
      isImmediateChildOfTopLevel: child.isImmediateChildOfTopLevel
    })));

    return (
      <div className={styles.childrenContainer} data-parent-id={parentId}>
        {uniqueChildren.map((child) => (
          <div key={child.Id} className={styles.childNode}>
            {this.renderEmployee(child, 1, false)}
          </div>
        ))}
      </div>
    );
  }

  public renderTopLevelExecutives(executives: IEmployee[]): JSX.Element {
    const uniqueExecutives = Array.from(new Map(executives.map(exec => [exec.Id, exec])).values());

    console.log('Rendering Top-Level Executives:', uniqueExecutives.map(exec => ({
      Id: exec.Id,
      Name: exec.EmployeeName?.Title,
      Children: exec.children?.map(child => child.Id) || [],
      isImmediateChildOfTopLevel: exec.isImmediateChildOfTopLevel
    })));

    return (
      <div className={styles.toplevelexecutives}>
        {uniqueExecutives.map((executive) => (
          <div key={executive.Id} className={styles.executiveColumn}>
            <div className={styles.managerNode}>
              {this.renderEmployee(executive, 0, false)}
            </div>
            {executive.children && executive.children.length > 0 && (
              this.renderChildrenHorizontally(executive.children, executive.Id)
            )}
          </div>
        ))}
      </div>
    );
  }

  public getParentChain(employeeId: number, employees: IEmployee[]): number[] {
    const chain: number[] = [];

    const findParentChain = (id: number, emps: IEmployee[]): boolean => {
      for (const emp of emps) {
        if (emp.Id === id) {
          chain.push(emp.Id);
          return true;
        }
        if (emp.children) {
          if (findParentChain(id, emp.children)) {
            chain.push(emp.Id);
            return true;
          }
        }
      }
      return false;
    };

    findParentChain(employeeId, employees);
    return chain;
  }

  public getChildrenIds(employee: IEmployee): number[] {
    const ids: number[] = [employee.Id];

    const collectChildIds = (emp: IEmployee) => {
      if (emp.children) {
        emp.children.forEach(child => {
          ids.push(child.Id);
          collectChildIds(child);
        });
      }
    };

    collectChildIds(employee);
    return ids;
  }

  public handleDirectorClick = (employee: IEmployee, event: React.MouseEvent) => {
    this.handleCardClick(employee, event);
  };

  public renderEmployee(employee: IEmployee, level: number, isNoDepartment: boolean = false, renderedIds: Set<number> = new Set()): JSX.Element {
    if (renderedIds.has(employee.Id)) {
      return <></>;
    }
    renderedIds.add(employee.Id);
  
    const hasChildren = employee.children && employee.children.length > 0;
    const isExpanded = this.state.expandedNodes.has(employee.Id);
    const employeeTitle = employee.Title || '';
    const isDirector = employeeTitle.toLowerCase().includes('director');
    const isSelected = this.state.selectedCard === employee.Id;
  
    const isTopLevelExecutive = this.state.allEmployees.some(exec => exec.Id === employee.Id);
    const isNoDepartmentFlag = employee.isNoDepartment || isNoDepartment;
  
    const cardBackgroundColor = employee.Color || '#dadbdc';
    const textColor = employee.FontColor || '#000000';
  
    return (
      <div className={styles.employeeSection}>
        <div
          className={`${this.getNodeClassName(employee)} ${isDirector ? styles.directorNode : ''} 
            ${isSelected ? styles.selected : ''} ${isExpanded ? styles.expanded : ''}`}
          onClick={(e) => this.handleCardClick(employee, e)}
          style={{ backgroundColor: cardBackgroundColor, color: textColor }}
        >
          <div className={styles.cardContent}>
            <div className={styles.employeeInfo}>
              <div className={styles.name}>{employee.EmployeeName?.Title || 'N/A'}</div>
              <div className={styles.title}>{employee.Title || 'N/A'}</div>
              {isTopLevelExecutive && !employee.EmployeeManagers && !employee.Department && (
                <div className={`${styles.department} ${styles.hidden}`}>N/A</div>
              )}
              {isNoDepartmentFlag && !isTopLevelExecutive && (
                <div className={`${styles.department} ${styles.hidden}`}>N/A</div>
              )}
              {!employee.Department && !isTopLevelExecutive && !isNoDepartmentFlag && (
                <div className={`${styles.department} ${styles.hidden}`}>N/A</div>
              )}
              {employee.Department && !isNoDepartmentFlag && (
                <div className={styles.department}>{employee.Department}</div>
              )}
            </div>
            <div className={styles.avatarContainer}>
              <div
                className={styles.avatar}
                style={{
                  backgroundImage: employee.EmployeeName?.PictureUrl 
                    ? `url('${employee.EmployeeName.PictureUrl}')` 
                    : `url('${this.props.siteurl}/SiteAssets/nouserimageicon.jpg')`,
                }}
              />
            </div>
          </div>
        </div>
  
        {hasChildren && isExpanded && !isTopLevelExecutive && (
          <div className={`${styles.departments} ${isDirector ? styles.directorChildren : ''}`}>
            <div className={`${styles.childrenWrapper} ${isDirector ? styles.verticalChildren : ''}`}>
              {employee.children!.map(child => (
                <div key={child.Id} className={styles.childContainer}>
                  {this.renderEmployee(child, level + 1, child.isNoDepartment, renderedIds)}
                </div>
              ))}
            </div>
          </div>
        )}
      </div>
    );
  }

  public collectHierarchyIds(employees: IEmployee[]): Set<number> {
    const ids = new Set<number>();

    const collectIds = (employee: IEmployee) => {
      ids.add(employee.Id);
      if (employee.children) {
        employee.children.forEach(child => collectIds(child));
      }
    };

    employees.forEach(emp => collectIds(emp));
    return ids;
  }

  public renderEmployeesByDepartment(employees: IEmployee[]): JSX.Element {
    const departments = new Map<string, IEmployee[]>();
    const topLevelExecutives: IEmployee[] = [];
    const seen = new Set<number>();

    employees.forEach(emp => {
      if (seen.has(emp.Id)) return;
      seen.add(emp.Id);

      if (emp.Department?.trim()) {
        const deptEmployees = departments.get(emp.Department) || [];
        deptEmployees.push(emp);
        departments.set(emp.Department, deptEmployees);
      } else if (!emp.EmployeeManagers?.length) {
        topLevelExecutives.push(emp);
      }
    });

    const buildDepartmentHierarchy = (deptEmployees: IEmployee[]): IEmployee[] => {
      const employeeMap = new Map<number, IEmployee>(deptEmployees.map(emp => [emp.Id, { ...emp, children: emp.children || [] }]));
      
      const deptExecutives: IEmployee[] = [];
      const seenInHierarchy = new Set<number>();

      deptEmployees.forEach(emp => {
        const employee = employeeMap.get(emp.Id);
        if (!employee || seenInHierarchy.has(employee.Id)) return;

        const isTopLevelInDept = !employee.EmployeeManagers?.length || !Array.from(employeeMap.values()).some(
          e => e.EmployeeName?.Title && employee.EmployeeManagers?.[0]?.Title === e.EmployeeName.Title
        );

        if (isTopLevelInDept) {
          deptExecutives.push(employee);
          seenInHierarchy.add(employee.Id);
        } else {
          const primaryManager = employee.EmployeeManagers![0];
          const manager = Array.from(employeeMap.values()).find(
            e => e.EmployeeName?.Title === primaryManager.Title
          );
          if (manager && !manager.children?.some(child => child.Id === employee.Id)) {
            if (!manager.children) {
              manager.children = [];
            }
            manager.children.push(employee);
            seenInHierarchy.add(employee.Id);
          }
        }
      });

      const result = deptExecutives.length ? deptExecutives : Array.from(employeeMap.values());
      return result;
    };

    const renderDepartmentHierarchy = (deptName: string, deptExecutives: IEmployee[]): JSX.Element => (
      <div key={deptName} className={styles.departmentColumn}>
        <div className={styles.departmentBox}>{deptName}</div>
        {deptExecutives.length ? (
          <div className={styles.toplevelexecutives}>
            {deptExecutives.map(executive => (
              <div key={executive.Id} className={styles.executiveColumn}>
                <div className={styles.managerNode}>
                  {this.renderEmployee(executive, 0, false)}
                </div>
                {executive.children?.length ? (
                  this.renderChildrenHorizontally(executive.children, executive.Id)
                ) : (
                  <div style={{ display: "none" }}>No children found for {executive.EmployeeName?.Title}</div>
                )}
              </div>
            ))}
          </div>
        ) : (
          <div>No top-level executives found for {deptName}</div>
        )}
      </div>
    );

    return (
      <div className={`${styles.departments} ${styles.selectedView}`}>
        {topLevelExecutives.length > 0 && this.renderTopLevelExecutives(topLevelExecutives)}
        {departments.size > 0 && (
          <div className={styles.childrenWrapperwithDepartment}>
            {Array.from(departments.entries()).map(([deptName, deptEmployees]) =>
              renderDepartmentHierarchy(deptName, buildDepartmentHierarchy(deptEmployees))
            )}
          </div>
        )}
      </div>
    );
  }

  public getViewNamesForOrgView(orgViewTitle: string): string[] {
    const { managerPlaceholders } = this.state;
    const viewNames = new Set<string>();

    managerPlaceholders.forEach(emp => {
      const hasOrgView = emp.Views && emp.Views.some(view => view.Title === orgViewTitle);
      if (hasOrgView && emp.ViewName) {
        if (Array.isArray(emp.ViewName)) {
          emp.ViewName.forEach(name => viewNames.add(name));
        } else if (typeof emp.ViewName === 'string') {
          const names = emp.ViewName.split(/[,;]/).map(name => name.trim());
          names.forEach(name => viewNames.add(name));
        }
      }
    });

    const uniqueViewNames = Array.from(viewNames);
    console.log(`Unique ViewNames for ${orgViewTitle}:`, uniqueViewNames);
    return uniqueViewNames;
  }

  public renderManagerTreeView(manager: IEmployee, viewNames: string[]): JSX.Element {
    let { selectedView, selectedOrgView } = this.state;
    if (!selectedView && selectedOrgView) {
      selectedView = selectedOrgView;
    }
    console.log("Current selectedView:", selectedView);
    return (
      <div className={styles.managerTree}>
        {selectedView && (
          <div className={styles.viewTitle}>
            {selectedView.toUpperCase()}
          </div>
        )}
        <div className={styles.managerCard}>
          {this.renderEmployee(manager, 0, false)}
        </div>

        {viewNames.length > 0 && (
          <div className={styles.treeConnector}>
            <div className={styles.verticalLine} />
            <div className={styles.horizontalLine} />
            <div className={styles.treeButtons}>
              {viewNames.map((viewName, index) => (
                <div key={index} className={styles.buttonWrapper}>
                  <div className={styles.buttonConnector} />
                  <button
                    className={styles.viewButton}
                    onClick={() => this.setState({ selectedView: viewName }, () => {
                      this.filterEmployeesBasedOnSearch();
                    })}
                  >
                    {viewName}
                  </button>
                </div>
              ))}
            </div>
          </div>
        )}
      </div>
    );
  }

  public renderScreen0(): JSX.Element {
    const { orgViews, loading, error, managerPlaceholders } = this.state;

    if (loading) {
      return <div>Loading views...</div>;
    }

    if (error) {
      return <div>Error: {error}</div>;
    }

    const hasViewNames = (orgViewTitle: string) => {
      const viewNames = this.getViewNamesForOrgView(orgViewTitle);
      return viewNames.length > 0;
    };

    const shouldRenderManagerTreeDirectly = (orgViewTitle: string) => {
      const matchingEmployees = managerPlaceholders.filter(emp => {
        const hasNoEmployeeName = !emp.EmployeeName;
        const hasManager = !!emp.EmployeeManagers && emp.EmployeeManagers.length > 0;
        const hasOrgView = emp.Views && emp.Views.some(view => view.Title === orgViewTitle);
        return hasNoEmployeeName && hasManager && hasOrgView && hasViewNames(orgViewTitle);
      });
      return matchingEmployees.length > 0;
    };

    return (
      <div className={styles.screen0Container}>
        <h2>Select a View</h2>
        <div className={styles.viewButtons}>
          {orgViews.length > 0 ? (
            orgViews.map(view => (
              <button
                key={view.Id}
                className={styles.viewButton}
                onClick={() => {
                  if (shouldRenderManagerTreeDirectly(view.Title)) {
                    this.setState({ currentScreen: 'Screen1', selectedOrgView: view.Title }, () => {
                      this.filterEmployeesBasedOnSearch();
                    });
                  } else {
                    const hasSubViews = hasViewNames(view.Title);
                    if (hasSubViews) {
                      this.setState({ currentScreen: 'Screen2', selectedOrgView: view.Title });
                    } else {
                      this.setState({ currentScreen: 'Screen1', selectedView: view.Title }, () => {
                        this.filterEmployeesBasedOnSearch();
                      });
                    }
                  }
                }}
              >
                {view.Title}
              </button>
            ))
          ) : (
            <div>No views available</div>
          )}
        </div>
      </div>
    );
  }

  public renderScreen2(): JSX.Element {
    const { selectedOrgView, loading, error } = this.state;

    if (loading) {
      return <div>Loading view names...</div>;
    }

    if (error) {
      return <div>Error: {error}</div>;
    }

    if (!selectedOrgView) {
      return <div>No OrgView selected</div>;
    }

    const viewNames = this.getViewNamesForOrgView(selectedOrgView);

    return (
      <div className={styles.screen0Container}>
        <h2>{selectedOrgView} - Select a Sub-View</h2>
        <div className={styles.viewButtons}>
          {viewNames.length > 0 ? (
            viewNames.map((viewName, index) => (
              <button
                key={index}
                className={styles.viewButton}
                onClick={() => this.setState({ currentScreen: 'Screen1', selectedView: viewName }, () => {
                  this.filterEmployeesBasedOnSearch();
                })}
              >
                {viewName}
              </button>
            ))
          ) : (
            <div>No sub-views available for {selectedOrgView}</div>
          )}
          <button
            className={styles.backButton}
            onClick={() => this.setState({ currentScreen: 'Screen0', selectedOrgView: null }, () => {
              this.filterEmployeesBasedOnCount();
            })}
          >
            Back to Views
          </button>
        </div>
      </div>
    );
  }

  public handleZoomIn = (): void => {
    this.setState(prevState => ({
      zoomLevel: Math.min(prevState.zoomLevel + 0.1, 2),
    }));
  };

  public handleZoomOut = (): void => {
    this.setState(prevState => ({
      zoomLevel: Math.max(prevState.zoomLevel - 0.1, 0.2),
    }));
  };

  public handleZoomDefault = (): void => {
    this.setState({ zoomLevel: 1 });
  };

  public renderScreen1(): JSX.Element {
    const { loading, error, employees, searchTerm, selectedView, selectedOrgView, managerPlaceholders, zoomLevel } = this.state;

    if (loading) {
      return <div>Loading organization chart...</div>;
    }

    if (error) {
      return <div>Error: {error}</div>;
    }

    const uniqueEmployees = Array.from(new Map(employees.map(emp => [emp.Id, emp])).values());

    const shouldRenderManagerTree = () => {
      if (!selectedOrgView || selectedView) return false;

      const matchingEmployees = managerPlaceholders.filter(emp => {
        const hasNoEmployeeName = !emp.EmployeeName;
        const hasManager = !!emp.EmployeeManagers && emp.EmployeeManagers.length > 0;
        const hasOrgView = emp.Views && emp.Views.some(view => view.Title === selectedOrgView);
        return hasNoEmployeeName && hasManager && hasOrgView;
      });

      return matchingEmployees.length > 0;
    };

    const viewClass = selectedView
      ? `view-${selectedView.toLowerCase().replace(/\s+/g, '-')}`
      : '';

    const dynamicClass = viewClass && styles[viewClass as keyof typeof styles] ? styles[viewClass as keyof typeof styles] : '';

    if (shouldRenderManagerTree()) {
      const matchingEmployees = managerPlaceholders.filter(emp => {
        const hasNoEmployeeName = !emp.EmployeeName;
        const hasManager = !!emp.EmployeeManagers && emp.EmployeeManagers.length > 0;
        const hasOrgView = emp.Views && emp.Views.some(view => view.Title === selectedOrgView);
        return hasNoEmployeeName && hasManager && hasOrgView;
      });

      if (matchingEmployees.length > 0) {
        const viewNames = new Set<string>();
        matchingEmployees.forEach(emp => {
          if (emp.ViewName) {
            if (Array.isArray(emp.ViewName)) {
              emp.ViewName.forEach(name => viewNames.add(name));
            } else if (typeof emp.ViewName === 'string') {
              emp.ViewName.split(/[,;]/).map(name => name.trim()).forEach(name => viewNames.add(name));
            }
          }
        });
        const uniqueViewNames = Array.from(viewNames);

        const managersMap = new Map<string, IEmployee>();
        matchingEmployees.forEach(emp => {
          emp.EmployeeManagers?.forEach(mgr => {
            const managerTitle = mgr.Title;
            if (managerTitle && !managersMap.has(managerTitle)) {
              const managerData = mgr;
              const manager: IEmployee = {
                Id: emp.Id,
                Title: emp.Title || 'N/A',
                EmployeeName: {
                  Title: managerData?.Title || 'N/A',
                  EMail: managerData?.EMail || 'N/A',
                  PictureUrl: managerData?.EMail
                    ? `${this.props.siteurl}/_layouts/15/userphoto.aspx?size=L&accountname=${managerData.EMail}`
                    : ''
                },
                EmployeeManagers: null,
                Department: emp.Department || null,
                Phone: emp.Phone || 0,
                Location: emp.Location || '',
                Views: emp.Views,
                ViewName: emp.ViewName,
                children: [],
              };
              managersMap.set(managerTitle, manager);
            }
          });
        });

        const uniqueManagers = Array.from(managersMap.values());

        return (
          <div className={`${styles.orgchartContainer} ${dynamicClass}`}>
            <div className={styles.searchBar}>
              <input
                type="text"
                placeholder="Search by..."
                value={searchTerm}
                onChange={this.handleSearchChange}
              />
              <button className={styles.zoomButton} onClick={this.handleZoomIn}>
                <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" fill="currentColor" className="bi bi-zoom-in" viewBox="0 0 16 16">
                  <path fillRule="evenodd" d="M6.5 12a5.5 5.5 0 1 0 0-11 5.5 5.5 0 0 0 0 11M13 6.5a6.5 6.5 0 1 1-13 0 6.5 6.5 0 0 1 13 0"/>
                  <path d="M10.344 11.742q.044.06.098.115l3.85 3.85a1 1 0 0 0 1.415-1.414l-3.85-3.85a1 1 0 0 0-.115-.1 6.5 6.5 0 0 1-1.398 1.4z"/>
                  <path fillRule="evenodd" d="M6.5 3a.5.5 0 0 1 .5.5V6h2.5a.5.5 0 0 1 0 1H7v2.5a.5.5 0 0 1-1 0V7H3.5a.5.5 0 0 1 0-1H6V3.5a.5.5 0 0 1 .5-.5"/>
                </svg>
              </button>
              <button className={styles.zoomButton} onClick={this.handleZoomOut}>
                <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" fill="currentColor" className="bi bi-zoom-out" viewBox="0 0 16 16">
                  <path fillRule="evenodd" d="M6.5 12a5.5 5.5 0 1 0 0-11 5.5 5.5 0 0 0 0 11M13 6.5a6.5 6.5 0 1 1-13 0 6.5 6.5 0 0 1 13 0"/>
                  <path d="M10.344 11.742q.044.06.098.115l3.85 3.85a1 1 0 0 0 1.415-1.414l-3.85-3.85a1 1 0 0 0-.115-.1 6.5 6.5 0 0 1-1.398 1.4z"/>
                  <path fillRule="evenodd" d="M3 6.5a.5.5 0 0 1 .5-.5h6a.5.5 0 0 1 0 1h-6a.5.5 0 0 1-.5-.5"/>
                </svg>
              </button>
              <button className={styles.zoomButton} onClick={this.handleZoomDefault}>
                <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" fill="currentColor" className="bi bi-arrows-fullscreen" viewBox="0 0 16 16">
                  <path fillRule="evenodd" d="M5.828 10.172a.5.5 0 0 0-.707 0l-4.096 4.096V11.5a.5.5 0 0 0-1 0v3.975a.5.5 0 0 0 .5.5H4.5a.5.5 0 0 0 0-1H1.732l4.096-4.096a.5.5 0 0 0 0-.707m4.344 0a.5.5 0 0 1 .707 0l4.096 4.096V11.5a.5.5 0 1 1 1 0v3.975a.5.5 0 0 1-.5.5H11.5a.5.5 0 0 1 0-1h2.768l-4.096-4.096a.5.5 0 0 1 0-.707m0-4.344a.5.5 0 0 0 .707 0l4.096-4.096V4.5a.5.5 0 1 0 1 0V.525a.5.5 0 0 0-.5-.5H11.5a.5.5 0 0 0 0 1h2.768l-4.096 4.096a.5.5 0 0 0 0 .707m-4.344 0a.5.5 0 0 1-.707 0L1.025 1.732V4.5a.5.5 0 0 1-1 0V.525a.5.5 0 0 1 .5-.5H4.5a.5.5 0 0 1 0 1H1.732l4.096 4.096a.5.5 0 0 1 0 .707"/>
                </svg>
              </button>
              <button
                className={styles.backButton}
                onClick={() => {
                  this.setState({ currentScreen: 'Screen0', searchTerm: '', selectedCard: null, employees: [...this.state.originalEmployees], selectedView: null, selectedOrgView: null }, () => {
                    this.filterEmployeesBasedOnCount();
                  });
                }}
              >
                Back to Views
              </button>
            </div>
            <div className={styles.orgchartWrapper} ref={this.orgchartWrapperRef}>
              <div className={styles.orgchart} style={{ transform: `scale(${zoomLevel})` }}>
                <div className={styles.scrollContainer}>
                  {uniqueViewNames.length > 0 && uniqueManagers.length > 0 ? (
                    uniqueManagers.map((manager, index) => (
                      <div key={index} className={styles.managerTreeWrapper}>
                        {this.renderManagerTreeView(manager, uniqueViewNames)}
                      </div>
                    ))
                  ) : (
                    <div>No sub-views or managers found</div>
                  )}
                </div>
              </div>
            </div>
          </div>
        );
      }
    }

    return (
      <div className={`${styles.orgchartContainer} ${dynamicClass}`}>
        <div className={styles.searchBar}>
          <input
            type="text"
            placeholder="Search by..."
            value={searchTerm}
            onChange={this.handleSearchChange}
          />
          <button className={styles.zoomButton} onClick={this.handleZoomIn}>
            <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" fill="currentColor" className="bi bi-zoom-in" viewBox="0 0 16 16">
              <path fillRule="evenodd" d="M6.5 12a5.5 5.5 0 1 0 0-11 5.5 5.5 0 0 0 0 11M13 6.5a6.5 6.5 0 1 1-13 0 6.5 6.5 0 0 1 13 0"/>
              <path d="M10.344 11.742q.044.06.098.115l3.85 3.85a1 1 0 0 0 1.415-1.414l-3.85-3.85a1 1 0 0 0-.115-.1 6.5 6.5 0 0 1-1.398 1.4z"/>
              <path fillRule="evenodd" d="M6.5 3a.5.5 0 0 1 .5.5V6h2.5a.5.5 0 0 1 0 1H7v2.5a.5.5 0 0 1-1 0V7H3.5a.5.5 0 0 1 0-1H6V3.5a.5.5 0 0 1 .5-.5"/>
            </svg>
          </button>
          <button className={styles.zoomButton} onClick={this.handleZoomOut}>
            <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" fill="currentColor" className="bi bi-zoom-out" viewBox="0 0 16 16">
              <path fillRule="evenodd" d="M6.5 12a5.5 5.5 0 1 0 0-11 5.5 5.5 0 0 0 0 11M13 6.5a6.5 6.5 0 1 1-13 0 6.5 6.5 0 0 1 13 0"/>
              <path d="M10.344 11.742q.044.06.098.115l3.85 3.85a1 1 0 0 0 1.415-1.414l-3.85-3.85a1 1 0 0 0-.115-.1 6.5 6.5 0 0 1-1.398 1.4z"/>
              <path fillRule="evenodd" d="M3 6.5a.5.5 0 0 1 .5-.5h6a.5.5 0 0 1 0 1h-6a.5.5 0 0 1-.5-.5"/>
            </svg>
          </button>
          <button className={styles.zoomButton} onClick={this.handleZoomDefault}>
            <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" fill="currentColor" className="bi bi-arrows-fullscreen" viewBox="0 0 16 16">
              <path fillRule="evenodd" d="M5.828 10.172a.5.5 0 0 0-.707 0l-4.096 4.096V11.5a.5.5 0 0 0-1 0v3.975a.5.5 0 0 0 .5.5H4.5a.5.5 0 0 0 0-1H1.732l4.096-4.096a.5.5 0 0 0 0-.707m4.344 0a.5.5 0 0 1 .707 0l4.096 4.096V11.5a.5.5 0 1 1 1 0v3.975a.5.5 0 0 1-.5.5H11.5a.5.5 0 0 1 0-1h2.768l-4.096-4.096a.5.5 0 0 1 0-.707m0-4.344a.5.5 0 0 0 .707 0l4.096-4.096V4.5a.5.5 0 1 0 1 0V.525a.5.5 0 0 0-.5-.5H11.5a.5.5 0 0 0 0 1h2.768l-4.096 4.096a.5.5 0 0 0 0 .707m-4.344 0a.5.5 0 0 1-.707 0L1.025 1.732V4.5a.5.5 0 0 1-1 0V.525a.5.5 0 0 1 .5-.5H4.5a.5.5 0 0 1 0 1H1.732l4.096 4.096a.5.5 0 0 1 0 .707"/>
            </svg>
          </button>
          <button
            className={styles.backButton}
            onClick={() => {
              if (this.state.selectedOrgView && selectedView) {
                this.setState({ currentScreen: 'Screen1', searchTerm: '', selectedCard: null, employees: [...this.state.originalEmployees], selectedView: null });
              } else if (this.state.selectedOrgView) {
                this.setState({ currentScreen: 'Screen2', searchTerm: '', selectedCard: null, employees: [...this.state.originalEmployees], selectedView: null });
              } else {
                this.setState({ currentScreen: 'Screen0', searchTerm: '', selectedCard: null, employees: [...this.state.originalEmployees], selectedView: null, selectedOrgView: null }, () => {
                  this.filterEmployeesBasedOnCount();
                });
              }
            }}
          >
            {this.state.selectedOrgView ? 'Back to Sub-Views' : 'Back to Views'}
          </button>
        </div>
        {selectedView && (
          <div className={styles.viewTitle}>
            {selectedView.toUpperCase()}
          </div>
        )}
        <div className={styles.orgchartWrapper} ref={this.orgchartWrapperRef}>
          <div className={styles.orgchart} style={{ transform: `scale(${zoomLevel})` }}>
            <div className={styles.scrollContainer}>
              {uniqueEmployees.length > 0 ? (
                this.renderEmployeesByDepartment(uniqueEmployees)
              ) : (
                <div>No employees found</div>
              )}
            </div>
          </div>
        </div>
        {this.renderPopup()}
      </div>
    );
  }

  public render(): React.ReactElement<IOrgchartProps> {
    const { currentScreen } = this.state;

    return (
      <div className={styles.orgchartMain}>
        {currentScreen === 'Screen0' && this.renderScreen0()}
        {currentScreen === 'Screen1' && this.renderScreen1()}
        {currentScreen === 'Screen2' && this.renderScreen2()}
      </div>
    );
  }
}