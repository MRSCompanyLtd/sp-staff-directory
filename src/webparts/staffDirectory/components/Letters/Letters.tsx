import * as React from "react";
import styles from "./Letters.module.scss";
import { ILettersProps } from "./ILettersProps";

const ALPHABET: string[] = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"];

const Letters: React.FC<ILettersProps> = ({ letter, onLetterClick }) => {
    const [selected, setSelected] = React.useState<string | undefined>(letter);

    const handleClick: React.MouseEventHandler<HTMLDivElement> = React.useCallback((e: React.MouseEvent<HTMLDivElement, MouseEvent>): void => {
        const ltr: string = e.currentTarget.innerText;
        
        if (ltr === selected) {
            setSelected(undefined);
            onLetterClick(undefined);
        } else {
            setSelected(ltr);
            onLetterClick(ltr);
        }
    }, [onLetterClick, selected]);

    React.useEffect(() => {
        setSelected(letter);
    }, [letter]);

    return (
        <div className={styles.letters}>
            {ALPHABET.map((item: string) => (
                <div
                    key={item}
                    style={{ fontWeight: 'bolder', flexGrow: 1 }}
                    onClick={handleClick}
                    className={`${styles.letter} ${selected === item && styles.active}`}>
                    {item}
                </div>
            ))}
        </div>
    );
}

export default Letters;
