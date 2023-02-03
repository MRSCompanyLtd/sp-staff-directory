import { SearchBox } from "office-ui-fabric-react/lib/SearchBox";
import * as React from "react";
import { ISearchProps } from "./ISearchProps";
import styles from "./Search.module.scss";

const Search: React.FC<ISearchProps> = ({ onChange, clear, value, submit }) => {
    return (
        <div className={styles.search}>
            <SearchBox
                className={styles.search}
                placeholder='Search first name, last name, department, or job title'
                onChange={onChange}
                onClear={clear}
                value={value}
                onSearch={submit}
            />
        </div>
    );
}

export default Search;