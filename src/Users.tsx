import React, { useState, useEffect } from 'react';
import { config } from './Config';
import { UserAgentApplication } from 'msal';
import { getUser, deleteUser, updateUser } from './GraphService';
import { Table, Button, Anchor, Drawer, Form, Input, Switch } from 'antd';
import 'antd/dist/antd.css';


const User = () => {

    const [users, setUser] = useState<any>([]);
    const { Link } = Anchor;
    const [drawerVisible, setDrawerVisible] = useState<boolean>(false)
    const [user, setOneUser] = useState<any>();




    useEffect(() => {
        getUserService().then(result => {
            console.log(result)
        });
    }, []);

    const columns = [
        {
            title: 'Display Name',
            dataIndex: 'displayName',
            key: 'displayName',
        },
        {
            title: 'Mail',
            dataIndex: 'mail',
            key: 'mail',
        },
        {
            title: '',
            dataIndex: 'id',
            key: 'id',
            render: (key: any, test: any) => <> <Button onClick={() => userUpdate(test)}>Update</Button>
                <Button onClick={() => { userDelete(key) }}>Delete</Button>
            </>,
        },
        // {
        //     title: 'Action',
        //     dataIndex: '',
        //     key: 'x',
        //     render: () => <a>Delete</a>
        // },
    ];

    const userAgentApplication = new UserAgentApplication({
        auth: {
            clientId: config.appId,
            redirectUri: config.redirectUri
        },
        cache: {
            cacheLocation: "sessionStorage",
            storeAuthStateInCookie: true
        }
    });

    const isInteractionRequired = (error: Error): boolean => {
        if (!error.message || error.message.length <= 0) {
            return false;
        }

        return (
            error.message.indexOf('consent_required') > -1 ||
            error.message.indexOf('interaction_required') > -1 ||
            error.message.indexOf('login_required') > -1
        );
    }

    const getAccessToken = async (scopes: string[]) => {

        try {
            var silentResult = await userAgentApplication.acquireTokenSilent({
                scopes: scopes
            });
            return silentResult.accessToken;
        } catch (err) {
            // If a silent request fails, it may be because the user needs
            // to login or grant consent to one or more of the requested scopes
            if (isInteractionRequired(err)) {
                var interactiveResult = await userAgentApplication.acquireTokenPopup({
                    scopes: scopes
                });

                return interactiveResult.accessToken;
            } else {
                throw err;
            }
        }
    }

    const getUserService = async () => {
        const accessToken = await getAccessToken(config.scopes);
        getUser(accessToken)
            .then((result: any) => {
                const { value } = result;
                setUser(value)
                console.log(value)
            });

    }

    const userDelete = async (id: any) => {
        const data = users.filter((user: any) => user.id !== id)
        setUser(data)
        const accessToken = await getAccessToken(config.scopes);
        deleteUser(accessToken, id)
            .then(res => console.log(res))
            .catch(res => console.log(res));
    }

    const userUpdate = (user: any) => {
        setDrawerVisible(true);
        setOneUser(user)

    }

    const onClose = () => {
        setDrawerVisible(false)
    }

    const userUpdateSave = async (user: any) => {
        
        const accessToken = await getAccessToken(config.scopes);
     //   setOneUser(user)
        console.log(user)

        updateUser(accessToken, user)
            .then(res => res)
    }

    return (
        <>
            {/* <Anchor>
            <Button><Link href='/adduser' title=''/>  </Button> 
            </Anchor> */}
            <Button href='/adduser'>Craete New User</Button>
            <br />
            <Table dataSource={users} columns={columns} />

            <Drawer
                title="Create a new account"
                width={720}
                onClose={onClose}
                visible={drawerVisible}
                bodyStyle={{ paddingBottom: 80 }}

            >
                <Form
                    onFinish={(values)=>userUpdateSave(values)}
                    initialValues={{
                        'displayName': user === undefined ? '' : (user.displayName !== null? user.displayName : ''),
                        'givenName': user === undefined ? '' : (user.givenName !== null? user.givenName : ''),
                        'id': user === undefined ? '' : (user.id !== null? user.id : ''),
                        'jobTitle': user === undefined ? '' : (user.jobTitle !== null? user.jobTitle : ''),
                        'mobilePhone': user === undefined ? '' : (user.mobilePhone !== null? user.mobilePhone : ''),
                        'officeLocation': user === undefined ? '' : (user.officeLocation !== null? user.officeLocation : ''),
                        'surname': user === undefined ? '' : (user.surname !== null? user.surname : ''),
                        'userPrincipalName': user === undefined ? '' : (user.userPrincipalName !== null? user.userPrincipalName : ''),
                        'mail': user === undefined ? '' : (user.mail!== null ? user.mail : ''),
                    }}
                >


                    <Form.Item
                        label="Display Name"
                        name="displayName"
                    >
                        <Input />
                    </Form.Item>
                    <Form.Item
                        label="Email"
                        name="mail"
                    >
                        <Input />
                    </Form.Item>
                    <br />
                    <Form.Item>
                        <Button type="primary" htmlType="submit">
                            update
                            </Button>
                        <Button onClick={onClose} type="default" htmlType="submit">
                            Cancel
                            </Button>
                    </Form.Item>
                </Form>
            </Drawer>

        </>
    );

}

export default User;