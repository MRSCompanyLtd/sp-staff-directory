import { Persona, PersonaSize, Spinner } from "office-ui-fabric-react";
import * as React from "react";
import { IPerson } from "../../IPerson";
import { IResultsProps } from "./IResultsProps";
import styles from "./Results.module.scss";
import LivePersona from "../Person/LivePersona";

const Results: React.FC<IResultsProps> = ({ people, loading, context }) => {
    return (
        <div className={styles.resultsContainer}>
            {loading ?
                <div className={styles.loading}>
                    <Spinner size={3} />
                </div>
            :
                people.length > 0 ?
                    <div className={styles.results}>
                        {people.map((item: IPerson) => (
                            <LivePersona
                                key={item.id}
                                template={
                                    <>
                                        <Persona
                                            text={item.displayName}
                                            secondaryText={item.jobTitle}
                                            tertiaryText={item.department}
                                            className={styles.result}
                                            size={PersonaSize.size72}
                                            imageUrl={item.picture}
                                            imageAlt={`${item.displayName} profile pic`}
                                            imageInitials={''}
                                            imageShouldFadeIn
                                        />
                                    </>
                                }
                                serviceScope={context.serviceScope}
                                upn={item.upn}
                            />
                        ))}
                    </div>
                :
                    <div className={styles.noResults}>
                        <span>No results for search query</span>
                    </div>
            }
        </div>
    );
}

export default Results;
