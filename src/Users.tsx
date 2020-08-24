import React, { useState, useEffect } from 'react';
import { config } from './Config';
import { UserAgentApplication } from 'msal';
import { getUser, deleteUser, updateUser } from './GraphService';
import { Table, Button, Drawer, Form, Input, Modal } from 'antd';
import 'antd/dist/antd.css';

const User = () => {

    const [users, setUser] = useState<any>([]);
    const [drawerVisible, setDrawerVisible] = useState<boolean>(false)
    const [user, setOneUser] = useState<any>();
    const { confirm } = Modal;

    useEffect(() => {
        getUserService().then(result => {
            console.log(result)
        });
    }, [user]);

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
            render: (key: any, test: any) => <>
                <Button style={{ 'marginRight': '3px' }} onClick={() => userUpdate(test)}>Update</Button>
                <Button onClick={() => { userDelete(key) }}>Delete</Button>
            </>,
        },

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

        confirm({
            title: 'Do you Want to delete these items?',
            // content: 'Some descriptions',
            async onOk() {
                console.log('OK');
                const data = users.filter((user: any) => user.id !== id)
                setUser(data)
                const accessToken = await getAccessToken(config.scopes);
                deleteUser(accessToken, id)
                    .then(res => console.log(res))
                    .catch(res => console.log(res));
            },
            onCancel() {
                console.log('Cancel');
            },
        });
    }

    const userUpdate = (user: any) => {
        setDrawerVisible(true);
        setOneUser(user)

    }


    const userUpdateSave = async (values: any) => {

        const accessToken = await getAccessToken(config.scopes);
        values.id = values.id !== undefined ? values.id : (user.id !== null ? user.id : '');
        values.displayName = values.displayName !== undefined || null ? values.displayName : (user.displayName !== null ? user.displayName : '');
        values.givenName = values.givenName !== undefined || null ? values.givenName : (user.givenName !== null ? user.givenName : null);
        values.jobTitle = values.jobTitle !== undefined || null ? values.jobTitle : (user.jobTitle !== null ? user.jobTitle : null);
        values.mobilePhone = values.mobilePhone !== undefined || null ? values.mobilePhone : (user.mobilePhone !== null ? user.mobilePhone : null);
        values.officeLocation = values.officeLocation !== undefined || null ? values.officeLocation : (user.officeLocation !== null ? user.officeLocation : null);
        values.surname = values.surname !== undefined || null ? values.surname : (user.surname !== null ? user.surname : null);
        values.userPrincipalName = values.userPrincipalName !== undefined || null ? values.userPrincipalName : (user.userPrincipalName !== null ? user.userPrincipalName : null);
        values.mail = values.mail !== undefined || null ? values.mail : (user.mail !== null ? user.mail : null);

        updateUser(accessToken, values)
            .then((res: any) => {
                setDrawerVisible(false);
                setOneUser(values)
            })
    }

    const onClose = () => {
        setDrawerVisible(false)
    }

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


    return (
        <>
            {/* <Anchor>
            <Button><Link href='/adduser' title=''/>  </Button> 
            </Anchor> */}
            <Button href='/adduser' >Create New User</Button>
            <br />
            <Table dataSource={users} columns={columns} style={{ 'paddingTop': '10px' }} />

            <Drawer
                title="Create a new account"
                width={720}
                onClose={onClose}
                visible={drawerVisible}
                bodyStyle={{ paddingBottom: 80 }}

            >
                <Form
                    onFinish={userUpdateSave}
                    initialValues={{
                        'displayName': user === undefined ? '' : (user.displayName !== null ? user.displayName : ''),
                        'givenName': user === undefined ? '' : (user.givenName !== null ? user.givenName : ''),
                        'id': user === undefined ? '' : (user.id !== null ? user.id : ''),
                        'jobTitle': user === undefined ? '' : (user.jobTitle !== null ? user.jobTitle : ''),
                        'mobilePhone': user === undefined ? '' : (user.mobilePhone !== null ? user.mobilePhone : ''),
                        'officeLocation': user === undefined ? '' : (user.officeLocation !== null ? user.officeLocation : ''),
                        'surname': user === undefined ? '' : (user.surname !== null ? user.surname : ''),
                        'userPrincipalName': user === undefined ? '' : (user.userPrincipalName !== null ? user.userPrincipalName : ''),
                        'mail': user === undefined ? '' : (user.mail !== null ? user.mail : ''),
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