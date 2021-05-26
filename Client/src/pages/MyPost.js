import React, { useContext } from 'react'
import {AuthContext} from '../context/auth';

import {Redirect} from 'react-router-dom';


import { FETCH_POSTS_QUERY } from "../util/graphql";
import { useQuery } from '@apollo/client';
import SinglePost from './SinglePost';



export default function MyPost() {

    const { user } = useContext(AuthContext);
    
    const { loading, error, data} = useQuery(FETCH_POSTS_QUERY);

    if(!loading) {

        const res = data.getPosts.filter(post => post.userid === user.id )
        return (
            console.log(res.id)
        )
    }
    else return <div>Loading...</div>;
}
