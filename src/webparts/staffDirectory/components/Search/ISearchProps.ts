export interface ISearchProps {
    onChange(e: React.ChangeEvent<HTMLInputElement>, newValue: string): void;
    clear(): void;
    value: string;
    submit(val: string): void;
}