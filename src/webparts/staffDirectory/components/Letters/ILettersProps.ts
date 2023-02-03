export interface ILettersProps {
    letter: string | undefined;
    onLetterClick(ltr: string): void;
    reset?(): void;
}