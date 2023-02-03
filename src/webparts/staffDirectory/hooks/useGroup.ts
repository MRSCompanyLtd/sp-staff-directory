import * as React from "react";
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { GraphFI, SPFx as grSPFx, graphfi, DefaultHeaders, IPagedResult } from '@pnp/graph';
import "@pnp/graph/users";
import "@pnp/graph/groups";
import "@pnp/graph/photos";
import "@pnp/graph/search";
import "@pnp/graph/batching";
import "@pnp/graph/members";
import { IPerson } from "../IPerson";
import { IMember } from "@pnp/graph/members";

interface IUseGroupReturn {
    getInitialLoad: (pageSize?: number) => Promise<IPerson[]>;
    searchPeople: (str: string, pageSize?: number) => Promise<IPerson[]>;
    searchLetter: (ltr: string, pageSize?: number) => Promise<IPerson[]>;
    getNextPage: (pages: number) => Promise<IPerson[]>;
    results: IPerson[];
    total: number;
}

const useGroup = (context: WebPartContext): IUseGroupReturn => {
    const [graph] = React.useState<GraphFI>(
        graphfi().using(grSPFx(context)).using(DefaultHeaders())
    );
    const [page, setPage] = React.useState<IPagedResult>();
    const [results, setResults] = React.useState<IPerson[]>([]);
    const [total, setTotal] = React.useState<number>(0);

    const cleanData = async (values: IMember[]): Promise<IPerson[]> => {
        try {
            const items: IPerson[] = await values
                .reduce(async (prev: Promise<IPerson[]>, curr: IMember) => {
                    // eslint-disable-next-line @typescript-eslint/no-explicit-any
                    const item: any = curr;
                    const photo: Blob = await graph
                        .users
                        .getById(item.id)
                        .photo
                        .getBlob()
                        .catch(() => null);
                    let photoUrl: string | null;

                    if (photo?.size) {
                        // eslint-disable-next-line @typescript-eslint/no-explicit-any
                        const url: any = window.URL || window.webkitURL;
                        photoUrl = url.createObjectURL(photo);
                    }

                    if (item.surname) {
                        (await prev).push({
                            id: item.id,
                            firstName: item.givenName,
                            lastName: item.surname,
                            displayName: item.displayName,
                            jobTitle: item.jobTitle,
                            email: item.mail,
                            businessPhone: item.businessPhones,
                            mobilePhone: item.mobilePhone,
                            department: item.department,
                            upn: item.userPrincipalName,
                            picture: photoUrl
                        });    
                    }
                    
                    return await prev;
            }, Promise.resolve([]));

            const ret: IPerson[] = items.sort((a: IPerson, b: IPerson) => a.lastName.localeCompare(b.lastName));

            return ret;
        }
        catch(e) {
            console.log(e);

            return [];
        }
    }

    const getInitialLoad = async (pageSize: number = 12): Promise<IPerson[]> => {
        try {
            const members: IMember[] = await graph.groups.getById('215d9254-08f3-4402-bbcb-5d99a5258aaa').members.top(pageSize)();
            const items: IPerson[] = await cleanData(members);

            setTotal(members.length);
            setResults(items);
            setPage(undefined);

            return items;
        }
        catch(e) {
            console.log(e);

            setPage(undefined);
            setResults([]);
            setTotal(0);

            return [];
        }
    }

    const searchPeople = async (query: string, pageSize: number = 12): Promise<IPerson[]> => {
        try {
            const search: IPagedResult = await graph
                .groups
                .getById('215d9254-08f3-4402-bbcb-5d99a5258aaa')
                .members
                .search(
                    `"displayName:${query}" OR "department:${query}" OR "jobTitle:${query}"
                `)
                .top(pageSize)
                .paged
                ();

            setPage(search);
            setTotal(search.count);
            
            const items: IPerson[] = await cleanData(search.value);

            setResults(items);
            setTotal(search.count);

            return items;
        }
        catch(e) {
            console.log(e);

            setPage(undefined);
            setResults([]);
            setTotal(0);

            return [];
        }
    }

    const getNextPage = async (pages: number = 1): Promise<IPerson[]> => {
        try {
            let newPage: IPagedResult = page;
            const all: IMember[] = [];

            for (let i: number = 0; i < pages; i++) {
                if (newPage.hasNext) {
                    newPage = await newPage.next();
                    all.push(...newPage.value);
                }                
            }

            const items: IPerson[] = await cleanData(all);

            const pastResults: IPerson[] = [...results];
            const newResults: IPerson[] = pastResults.concat(items);

            setResults(newResults);
            setPage(newPage);

            return items;
        }
        catch(e) {
            console.log(e);

            setPage(undefined);
            setResults([]);
            setTotal(0);

            return [];
        }
    }

    const searchLetter = async (ltr: string, pageSize: number = 12): Promise<IPerson[]> => {
        try {
            const search: IPagedResult = await graph
                .groups
                .getById('215d9254-08f3-4402-bbcb-5d99a5258aaa')
                .members
                .search(`"surname:${ltr}"`)
                .top(pageSize)
                .paged
                ();

            const items: IPerson[] = await cleanData(search.value);

            setResults(items);
            setPage(search);
            setTotal(search.count);

            return items;
        }
        catch(e) {
            console.log(e);

            setPage(undefined);
            setResults([]);
            setTotal(0);

            return [];
        }
    }

    return { searchPeople, getInitialLoad, getNextPage, searchLetter, total, results };
}

export default useGroup;
