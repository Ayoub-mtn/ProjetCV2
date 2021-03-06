import React, { useContext, useRef, useState } from "react";
import "./post.css";
import { AuthContext } from "../context/auth";
import DeleteButton from "../components/DeleteButton";
import { ADD_EDUCATION, ADD_Work, FETCH_POST_QUERY } from "../util/graphql";
import { useMutation, useQuery } from "@apollo/client";
import { Button, Card, Image } from "semantic-ui-react";
import AddExtra from "../components/AddWork";
import AddEducation from "../components/AddEducation";



function SinglePost(props) {
  const postId = props.match.params.postId;
  const { user } = useContext(AuthContext);


  const { data, loading } = useQuery(FETCH_POST_QUERY, {
    variables: {
      postId,
    },
  });
  
  function deletePostCallback() {
    props.history.push("/");
  }
  let postMarkup;
  if (loading) return <p>Loading...</p>;
  else {
    const {
      id,
      userid,
      firstname,
      lastname,
      job,
      phone,
      address,
      workExperience,
      education,
    } = data.getPost;
    //const postKeys = Object.values(getPost.education);

    postMarkup = (
      <div class="wrapper clearfix">
        <div class="left">
          <div class="name-hero">
            <div class="me-img"></div>
            <Image
          floated="right"
          size="small"
          src="https://react.semantic-ui.com/images/avatar/large/molly.png"
        />
            <div class="name-text">
              <h1>
                {firstname} <em>{lastname}</em>
              </h1>
              <p>{address}</p>
              <p>{user.email}n</p>
              <p>{phone}</p>
            </div>
          </div>
        </div>
        <div class="right">
          <div class="inner">
            <section>
              <h1>Work Experience</h1>
              {user && (
               <AddExtra data={data.getPost} />
              )}
            </section>
            <section>
              <h1>Education</h1>
              {user && (
                <AddEducation data={data.getPost} />
                )}
            </section>
            <section>
              <h1>Technical Skills</h1>
              <ul class="skill-set">
                <li>To be added</li>
              </ul>
            </section>
            <section>
              <h1>References</h1>
              <p>To be Added</p>
            </section>
            <section>
              <h1>Hobbies</h1>
              <ul class="skill-set">
                <li>Faith</li>
                <li>Biblical Studies</li>
                <li>Playing Guitar</li>
                <li>Song Writing</li>
                <li>Health & Nutrition</li>
                <li>Reading</li>
              </ul>
            </section>
          </div>
        </div>
      </div>
    );
  }
  return postMarkup;
}
export default SinglePost;
