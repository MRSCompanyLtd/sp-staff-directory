import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IPersonProps } from "../Person/IPersonProps";

export interface IResultsProps {
    people: IPersonProps[];
    loading: boolean;
    context: WebPartContext;
}