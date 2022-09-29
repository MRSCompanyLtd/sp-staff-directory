import * as React from "react";
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { GraphFI, SPFx as grSPFx, graphfi, DefaultHeaders, IPagedResult } from '@pnp/graph';
import "@pnp/graph/users";
import "@pnp/graph/photos";
import "@pnp/graph/search";
import "@pnp/graph/batching";
import { IPerson } from "../IPerson";
import { IUser, IUsers } from "@pnp/graph/users";

interface IUseSearchReturn {
    getInitialLoad: (pageSize?: number) => Promise<IPerson[]>;
    searchPeople: (str: string, pageSize?: number) => Promise<{ items: IPerson[], total: number }>;
    searchLetter: (ltr: string) => Promise<{ items: IPerson[], total: number }>;
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

    // React.useEffect(() => {
    //     Promise.resolve(graph.users.filter(`userType ne 'guest'`).count().then((res: number) => {
    //         setTotal(res);
    //     })).catch(() => setTotal(0));

    //     // eslint-disable-next-line react-hooks/exhaustive-deps
    // }, []);

    async function getPhoto(id: string): Promise<Blob> {
        return await graph.users.getById(id).photo.getBlob().catch(() => null);
    }

    async function getInitialLoad(pageSize: number = 12): Promise<IPerson[]> {
        try {
            const items: IPerson[] = [];

            const res = await graph.me.people
                .top(pageSize)
                .filter("personType.class eq 'Person'")
                .select(
                    "id", "displayName", "givenName", "surname", "jobTitle", "department",
                    "scoredEmailAddresses", "userPrincipalName", "phones"
                )
                ();

            for (let i: number = 0; i < res.length; i++) {
                // eslint-disable-next-line @typescript-eslint/no-explicit-any
                const item: any = res[i];
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
                    email: item.scoredEmailAddresses[0].address,
                    businessPhone: item.phones.find((p: { type: string, number: string}) => p.type === 'business')?.number ?? null,
                    mobilePhone: item.phones.find((p: { type: string, number: string}) => p.type === 'mobile')?.number ?? null,
                    department: item.department,
                    upn: item.userPrincipalName,
                    picture: photoUrl
                });
            }

            setResults(items);
            setTotal(12);

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
            const newResults: IPerson[] = pastResults.concat(items).sort((a: IPerson, b: IPerson) => a.lastName.localeCompare(b.lastName));

            setResults(newResults);
            setPage(newPage);

            return items;
        }
        catch(e) {
            console.log(e);

            return [];
        }
    }

    async function searchPeople(str: string, pageSize: number = 12): Promise<{ items: IPerson[], total: number }> {
        try {
            if (str === '') {
                const total: number = await graph.users.filter(`userType ne 'guest'`).count();
                const search: IPagedResult = await graph.users
                    .top(pageSize)
                    .select(
                        'id', 'department', 'displayName', 'givenName', 'surname', 'jobTitle',
                        'mail', 'businessPhones', 'mobilePhone', 'userPrincipalName'
                    )
                    .paged
                    ();

                const all: IUser[] = [];
                const items: IPerson[] = [];

                all.push(...search.value);

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
                        mobilePhone: item.mobilePhone,
                        department: item.department,
                        upn: item.userPrincipalName,
                        picture: photoUrl
                    });
                }

                const ret: IPerson[] = items.sort((a: IPerson, b: IPerson) => a.lastName.localeCompare(b.lastName));

                setResults(ret);
                setTotal(total);
                setPage(search);

                return { total, items };
            } else {
                const total: number = await graph.users.search(`("displayName:${str}" OR "department:${str}" OR "jobTitle:${str}")`).count();
                const search: IUsers = await graph
                    .users
                    .search(`("displayName:${str}" OR "department:${str}" OR "jobTitle:${str}")`)
                    .select(
                        'id', 'department', 'displayName', 'givenName', 'surname', 'jobTitle',
                        'mail', 'businessPhones', 'mobilePhone', 'userPrincipalName'
                    )
                    ();

                const items: IPerson[] = [];

                for (let i: number = 0; i < search.length; i++) {
                    // eslint-disable-next-line @typescript-eslint/no-explicit-any
                    const item: any = search[i];
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
                    })
                }

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

    async function searchLetter(ltr: string, pageSize: number = 12): Promise<{ items: IPerson[], total: number }> {
        try {
            const total: number = await graph.users
                .filter(`startsWith(givenname, '${ltr}') or startsWith(surname, '${ltr}') and userType ne 'guest'`)
                .count
                ();

            const search: IPagedResult = await graph
                .users
                .filter(`startsWith(givenname, '${ltr}') or startsWith(surname, '${ltr}')`)
                .top(pageSize)
                .select(
                    'id', 'department', 'displayName', 'givenName', 'surname', 'jobTitle',
                    'mail', 'businessPhones', 'mobilePhone', 'userPrincipalName'
                )
                .paged
                ();

            setPage(search);
            setTotal(total);

            const items: IPerson[] = [];

            for (let i: number = 0; i < search.value.length; i++) {
                // eslint-disable-next-line @typescript-eslint/no-explicit-any
                const item: any = search.value[i];
                const photo: Blob = await graph.users.getById(item.id).photo.getBlob().catch(() => null);
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
