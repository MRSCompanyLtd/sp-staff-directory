import * as React from "react";
import { Persona } from "office-ui-fabric-react";
import { IPersonProps } from "./IPersonProps";
import styles from "./Person.module.scss";

const Person: React.FC<IPersonProps> = ({
    id,
    department,
    displayName,
    email,
    jobTitle,
    phone,
    photo
}) => {
    return (
        <div className={styles.person}>
            <Persona
                imageUrl={photo}
                imageInitials={displayName[0] + displayName[displayName.indexOf(' ')]}
                text={displayName}
                secondaryText={jobTitle}
                tertiaryText={phone}
            />
        </div>
    );
}

export default Person;
