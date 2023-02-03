import * as React from "react";
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { GraphFI, SPFx as grSPFx, graphfi, DefaultHeaders, IPagedResult } from '@pnp/graph';
import "@pnp/graph/users";
import "@pnp/graph/groups";
import "@pnp/graph/photos";
import "@pnp/graph/search";
import "@pnp/graph/batching";
import { IPerson } from "../IPerson";
import { IPeople, IUser } from "@pnp/graph/users";

interface IUseSearchReturn {
    getInitialLoad: (pageSize?: number, query?: string) => Promise<IPerson[]>;
    searchPeople: (str: string, pageSize?: number, query?: string) => Promise<{ items: IPerson[], total: number }>;
    searchLetter: (ltr: string, pageSize?: number, query?: string) => Promise<{ items: IPerson[], total: number }>;
    getNextPage: (pages: number) => Promise<IPerson[]>;
    results: IPerson[];
    total: number;
}

const useSearch = (context: WebPartContext): IUseSearchReturn => {
    const [graph] = React.useState<GraphFI>(
        graphfi().using(grSPFx(context)).using(DefaultHeaders())
    );
    const [page, setPage] = React.useState<IPagedResult>();
    const [results, setResults] = React.useState<IPerson[]>([]);
    const [total, setTotal] = React.useState<number>(0);

    async function getPhoto(id: string): Promise<Blob> {
        return await graph.users.getById(id).photo.getBlob().catch(() => null);
    }

    async function getInitialLoad(pageSize: number = 12, query: string | undefined = ''): Promise<IPerson[]> {
        try {
            const selectFields: string[] = [
                'id', 'displayName', 'givenName', 'surname', 'jobTitle', 'department',
                'scoredEmailAddresses', 'userPrincipalName', 'phones'
            ]

            let filterQuery: string = `(personType/class eq 'Person' and surname ne null)`

            if (query !== undefined && query !== '') {
                filterQuery = `${filterQuery} and ${query}`;
            }

            const res: IPeople[] = await graph.me.people
                .filter(filterQuery)
                .top(pageSize)
                .select(selectFields.join(','))
                ();

            const items: IPerson[] = await res.reduce(async (prev: Promise<IPerson[]>, curr: IPeople) => {
                // eslint-disable-next-line @typescript-eslint/no-explicit-any
                const item: any = curr;
                const photo: Blob = await getPhoto(item.id);
                let photoUrl: string | null;

                if (photo?.size) {
                    // eslint-disable-next-line @typescript-eslint/no-explicit-any
                    const url: any = window.URL || window.webkitURL;
                    photoUrl = url.createObjectURL(photo);
                }

                (await prev).push({
                    id: item.id,
                    firstName: item.givenName,
                    lastName: item.surname,
                    displayName: item.displayName,
                    jobTitle: item.jobTitle,
                    email: item.scoredEmailAddresses[0].address,
                    businessPhone: item.phones.find((p: { type: string, number: string}) => p.type === 'business')?.number ?? null,
                    mobilePhone: item.phones.find((p: { type: string, number: string}) => p.type === 'mobile')?.number ?? null,
                    department: item.department,
                    upn: item.userPrincipalName,
                    picture: photoUrl
                });

                return await prev;
            }, Promise.resolve([]));

            setResults(items);
            setTotal(items.length);

            return items;
        }
        catch(e) {
            console.log(e);

            setResults([]);

            return [];
        }
    }

    async function getNextPage(pages: number = 1): Promise<IPerson[]> {
        try {
            const items: IPerson[] = [];
            const all: IUser[] = [];
            let newPage: IPagedResult = page;

            for (let i: number = 0; i < pages; i++) {
                if (newPage.hasNext) {
                    newPage = await newPage.next();
                    all.push(...newPage.value);
                }
            }

            for (let i: number = 0; i < all.length; i++) {
                // eslint-disable-next-line @typescript-eslint/no-explicit-any
                const item: any = all[i];
                const photo: Blob = await getPhoto(item.id);
                let photoUrl: string | null;
                if (photo?.size) {
                    // eslint-disable-next-line @typescript-eslint/no-explicit-any
                    const url: any = window.URL || window.webkitURL;
                    photoUrl = url.createObjectURL(photo);
                }

                items.push({
                    id: item.id,
                    firstName: item.givenName,
                    lastName: item.surname,
                    displayName: item.displayName,
                    jobTitle: item.jobTitle,
                    email: item.mail,
                    businessPhone: item.businessPhones,
                    department: item.department,
                    mobilePhone: item.mobilePhone,
                    upn: item.userPrincipalName,
                    picture: photoUrl
                });
            }

            const pastResults: IPerson[] = [...results];
            const newResults: IPerson[] = pastResults.concat(items).sort((a: IPerson, b: IPerson) => a.lastName ? a.lastName.localeCompare(b.lastName) : 1);

            setResults(newResults);
            setPage(newPage);

            return items;
        }
        catch(e) {
            console.log(e);

            return [];
        }
    }

    async function searchPeople(str: string, pageSize: number = 12, query: string | undefined = ''): Promise<{ items: IPerson[], total: number }> {
        try {
            let filterQuery: string = `(userType ne 'Guest' and accountEnabled eq true and surname ne null)`;
            if (query !== undefined && query !== '') {
                filterQuery = `${filterQuery} and ${query}`;
            }

            if (str === '') {
                const items: IPerson[] = await getInitialLoad();

                return { total: pageSize, items };
            } else {
                const total: number = await graph.users.filter(filterQuery).search(`("displayName:${str}" OR "department:${str}" OR "jobTitle:${str}")`).count();
                const search: IUser[] = await graph
                    .users
                    .filter(filterQuery)
                    .search(`("displayName:${str}" OR "department:${str}" OR "jobTitle:${str}")`)
                    .select(
                        'id', 'department', 'displayName', 'givenName', 'surname', 'jobTitle',
                        'mail', 'businessPhones', 'mobilePhone', 'userPrincipalName'
                    )
                    ();

                    const items: IPerson[] = await search.reduce(async (prev: Promise<IPerson[]>, curr: IUser) => {
                        // eslint-disable-next-line @typescript-eslint/no-explicit-any
                        const item: any = curr;
                        const photo: Blob = await getPhoto(item.id);
                        let photoUrl: string | null;
    
                        if (photo?.size) {
                            const url: URL | typeof webkitURL = window.URL || window.webkitURL;
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

                setResults(ret);
                setTotal(total);

                return { items, total };
            }
        }
        catch(e) {
            console.log(e);

            return { items: [], total: 0 };
        }
    }

    async function searchLetter(ltr: string, pageSize: number = 12, query: string | undefined = ''): Promise<{ items: IPerson[], total: number }> {
        try {
            let filterQuery: string = `(startsWith(givenname, '${ltr}') or (surname ne null and startsWith(surname, '${ltr}')) and userType eq 'Member' and accountEnabled eq true)`
            if (query !== undefined && query !== '') {
                filterQuery = `${filterQuery} and ${query}`;
            }

            const total: number = await graph.users
                .filter(filterQuery)
                .count
                ();

            const search: IPagedResult = await graph
                .users
                .filter(filterQuery)
                .top(pageSize)
                .select(
                    'id', 'department', 'displayName', 'givenName', 'surname', 'jobTitle',
                    'mail', 'businessPhones', 'mobilePhone', 'userPrincipalName'
                )
                .paged
                ();

            setPage(search);
            setTotal(total);

            const items: IPerson[] = await search.value.reduce(async (prev: Promise<IPerson[]>, curr: IUser) => {
                // eslint-disable-next-line @typescript-eslint/no-explicit-any
                const item: any = curr;
                const photo: Blob = await getPhoto(item.id);
                let photoUrl: string | null;

                if (photo?.size) {
                    const url: URL | typeof webkitURL = window.URL || window.webkitURL;
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

            setResults(ret);

            return { items, total };
        }
        catch(e) {
            console.log(e);

            return { items: [], total: 0 };
        }
    }

    return { searchPeople, searchLetter, getNextPage, getInitialLoad, results, total }
}

export default useSearch;
