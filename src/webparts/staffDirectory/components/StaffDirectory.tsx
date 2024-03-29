import * as React from 'react';
import styles from "./StaffDirectory.module.scss";
import { IStaffDirectoryProps } from './IStaffDirectoryProps';
import Letters from './Letters/Letters';
import Search from './Search/Search';
import Results from './Results/Results';
import { Dropdown, IDropdownOption } from 'office-ui-fabric-react';
import useGroup from '../hooks/useGroup';

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
  customQuery
}) => {
  const [loading, setLoading] = React.useState<boolean>(true);
  const [initial, setInitial] = React.useState<boolean>(false);
  const [search, setSearch] = React.useState<string>('');
  const [page, setPage] = React.useState<number>(1);
  const [letter, setLetter] = React.useState<string | undefined>(undefined);
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

  const { searchPeople, getNextPage, getInitialLoad, searchLetter, results, total } = useGroup(context);

  const get: () => Promise<void> = React.useCallback(async () => {
    setLoading(true);
    setSearch('');
    await getInitialLoad(pageSize);
  }, [pageSize, getInitialLoad]);

  React.useEffect(() => {
    if (initial) {
      Promise.resolve(get().then(() => {
        setPage(1);
        setLoading(false);
      })).catch(() => {
        setLoading(false);
      });      
    } else {
      Promise.resolve(getInitialLoad(pageSize).then(() => {
        setInitial(true);
        setLoading(false);
      })).catch(() => {
        setLoading(false);
      });  
    }

    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [pageSize, customQuery]);

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

    if (ltr === '' || ltr === undefined) {
      setSelectedDept('');
      await getInitialLoad(pageSize);
      setLetter(ltr);
      setPage(1);
    } else {
      setSelectedDept('');
      await searchLetter(ltr, pageSize);
      setLetter(ltr);
      setPage(1);
    }

    setLoading(false);
  }, [getInitialLoad, searchLetter, pageSize]);

  const handleSearchChange: (event?: React.ChangeEvent<HTMLInputElement>, newValue?: string) => void = React.useCallback((event?: React.ChangeEvent<HTMLInputElement>, newValue?: string): void => {
    setSearch(newValue);
  }, []);

  const handleSearch: (val: string) => Promise<void> = React.useCallback(async (val: string) => {
    setLoading(true);
    setLetter(undefined);
    setSelectedDept('');
    if (val === '' || val === undefined) {
      await getInitialLoad(pageSize);
    } else {
      await searchPeople(val, pageSize);  
    }
    setPage(1);
    setLoading(false);
  }, [searchPeople, getInitialLoad, pageSize]);

  const clearSearch: () => Promise<void> = React.useCallback(async () => {
    setLoading(true);
    setLetter(undefined);
    setSelectedDept('');
    await getInitialLoad(pageSize);
    setSearch('');
    setPage(1);
    setLoading(false);
  }, [getInitialLoad, pageSize]);

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
    await searchPeople(dept, pageSize);
  }, [searchPeople, pageSize]);

  return (
    <section className={`${styles.staffDirectory} ${hasTeamsContext ? styles.teams : ''}`}>
      <div className={styles.app}>
      {title && title !== '' &&
      <div className={styles.title}>
        <h2>{title}</h2>
      </div>
      }
      <div className={styles.searchArea}>
        <Letters letter={letter} onLetterClick={onLetterClick} />
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
      <Results people={
        results.slice((page - 1) * pageSize, (page * pageSize < total ? page * pageSize : total))
        } loading={loading} context={context}
      />
      {pages(page)}
      </div>
    </section>
  );
}

export default StaffDirectory;
