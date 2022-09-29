import * as React from 'react';
import styles from "./StaffDirectory.module.scss";
import { IStaffDirectoryProps } from './IStaffDirectoryProps';
import Letters from './Letters/Letters';
import Search from './Search/Search';
import Results from './Results/Results';
import useSearch from '../hooks/useSearch';
import { Dropdown, IDropdownOption } from 'office-ui-fabric-react';

interface IPropertyFieldCollectionDepartments {
  departmentKey: string;
  departmentName: string;
  uniqueId: string;
  sortIdx: number;
}

const StaffDirectory: React.FC<IStaffDirectoryProps> = ({
  title,
  pageSize,
  departments,
  showDepartmentFilter,
  context,
  hasTeamsContext,
  userDisplayName
}) => {
  const [loading, setLoading] = React.useState<boolean>(true);
  const [search, setSearch] = React.useState<string>('');
  const [page, setPage] = React.useState<number>(1);
  const [dropdownOptions, setDropdownOptions] = React.useState<IDropdownOption[]>(
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    (departments as any[]).reduce((prev: IDropdownOption[], curr: IPropertyFieldCollectionDepartments): IDropdownOption[] => {
      prev.push({
        key: curr.departmentKey,
        text: curr.departmentName
      });

      return prev;
    }, [{ key: '', text: 'All departments' }])
  );
  const [selectedDept, setSelectedDept] = React.useState<string>('');

  const { searchPeople, searchLetter, getNextPage, getInitialLoad, results, total } = useSearch(context);

  const get: () => Promise<void> = React.useCallback(async () => {
    setLoading(true);
    setSearch('');
    await searchPeople('', pageSize);
  }, [pageSize, searchPeople]);

  React.useEffect(() => {
    Promise.resolve(getInitialLoad(pageSize).then(() => {
      setLoading(false);
    })).catch(() => {
      setLoading(false);
    });

    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  React.useEffect(() => {
    Promise.resolve(get().then(() => {
      setPage(1);
      setLoading(false);
    })).catch(() => {
      setLoading(false);
    });

    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [pageSize]);

  React.useEffect(() => {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    setDropdownOptions((departments as any[]).reduce((prev: IDropdownOption[], curr: IPropertyFieldCollectionDepartments) => {
      prev.push({
        key: curr.departmentKey,
        text: curr.departmentName
      });

      return prev;
    }, [{ key: '', text: 'All departments' }]));
  }, [departments]);

  const onLetterClick: (ltr: string) => Promise<void> = React.useCallback(async (ltr: string): Promise<void> => {
    setLoading(true);

    if (ltr === '') {
      await searchPeople('');
      setPage(1);
    } else {
      await searchLetter(ltr);
      setPage(1);
    }

    setLoading(false);
  }, [searchPeople, searchLetter]);

  const handleSearchChange: (event?: React.ChangeEvent<HTMLInputElement>, newValue?: string) => void = React.useCallback((event?: React.ChangeEvent<HTMLInputElement>, newValue?: string): void => {
    setSearch(newValue);
  }, []);

  const handleSearch: (val: string) => Promise<void> = React.useCallback(async (val: string) => {
    setLoading(true);
    await searchPeople(val);
    setPage(1);
    setSelectedDept('');
    setLoading(false);
  }, [searchPeople]);

  const clearSearch: () => Promise<void> = React.useCallback(async () => {
    setLoading(true);
    await searchPeople('');
    setSearch('');
    setSelectedDept('');
    setPage(1);
    setLoading(false);
  }, [searchPeople]);

  const nextPage: (e: React.MouseEvent<HTMLDivElement, MouseEvent>) => Promise<void> = React.useCallback(async (e: React.MouseEvent<HTMLDivElement, MouseEvent>) => {
    const newVal: number = Number(e.currentTarget.innerText);
    const oldVal: number = page;
    setPage(newVal);
    if (newVal > oldVal) {
      setLoading(true);
      const num: number = newVal - oldVal;
      await getNextPage(num);
      setLoading(false);
    }
  }, [getNextPage, page]);

  const pages: (page: number) => JSX.Element = React.useCallback((page: number): JSX.Element => {
    const amount: number = Math.ceil(total / pageSize);
    const elements: JSX.Element[] = [];

    for (let i: number = 1; i <= amount; i++) {
      elements.push(
        <div
          key={i}
          className={`${styles.page} ${page === i && styles.active}`}
          onClick={nextPage}
        >
          {i}
        </div>
      );
    }

    return (
      <div className={styles.paging}>
        {elements}
      </div>
    );
  }, [nextPage, total, pageSize]);

  const handleSelect: (event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption, index?: number) => Promise<void> = React.useCallback(async (e, option) => {
    const dept: string = option.key.toString();
    setSelectedDept(dept);
    setSearch(dept);
    await searchPeople(dept);
  }, [searchPeople]);

  return (
    <section className={`${styles.staffDirectory} ${hasTeamsContext ? styles.teams : ''}`}>
      <div className={styles.app}>
      {title && title !== '' &&
      <div className={styles.title}>
        <h2>{title}</h2>
      </div>
      }
      <div className={styles.searchArea}>
        <Letters onLetterClick={onLetterClick} />
        {showDepartmentFilter &&
          <Dropdown
            options={dropdownOptions}
            onChange={handleSelect}
            selectedKey={selectedDept}
            styles={{ root: { width: '100%', marginBottom: '8px' } }}
          />
        }
        <Search onChange={handleSearchChange} clear={clearSearch} value={search} submit={handleSearch} />        
      </div>
      <Results people={results.slice((page - 1) * pageSize, page * pageSize)} loading={loading} context={context} />
      {pages(page)}
      </div>
    </section>
  );
}

export default StaffDirectory;
