import Sudoku from "./Sudoku/Sudoku";
import VBA_PT from "./VBA Pivot Tables/vba-PT";
import "./projects.scss";
import { useState } from "react";
import ArrowBackIosIcon from '@mui/icons-material/ArrowBackIos';
import ArrowForwardIosIcon from '@mui/icons-material/ArrowForwardIos';
import LanguageOutlinedIcon from '@mui/icons-material/LanguageOutlined';
import GitHubIcon from '@mui/icons-material/GitHub';

export default function Projects() {

    const [currentSlide,setCurrentSlide] = useState(0);

    const data = [
        {
            id: "1",
            icon: "assets/globe.png",
            title: "Sudoku",
            desc: "VBA based sudoku game",
            img: "assets/sudoku.png",
            livesite: "http://cj8355.github.io/RetroLand",
            repo: "https://github.com/cj8355/RetroLand",
            vid: "assets/sudoku-vid.webm",
            techUsed: [ "React", "Styled Components", "Firebase", "Material UI"]
        },
        {
            id: "2",
            icon: "assets/globe.png",
            title: "VBA Pivot Tables",
            desc: "Automatically create Pivot Tables using VBA",
            img: "assets/vba.png",
            livesite: "http://cj8355.github.io/RetroLand",
            repo: "https://github.com/cj8355/RetroLand",
            vid: "assets/vba-vid.webm",
            techUsed: [ "React", "Styled Components", "Firebase", "Material UI"]
        },
       
    ];
    
    const handleClick = (way)=> {
        way === "left" ? setCurrentSlide(currentSlide > 0 ? currentSlide - 1 : 2) :
        setCurrentSlide(currentSlide < data.length - 1 ? currentSlide + 1 : 0);
    }

    return (
        <div className="Container">


            <div className="slider" style={{transform: `translateX(-${currentSlide * 100}vw)` }}>
                {data.map((d) => (
                <div className="container" key={d.id}>
                    <div className="item">
                        <div className="left">
                            <div className="leftContainer">
                                {/* <div className="imgContainer">
                                    <img src={process.env.PUBLIC_URL + "/" +  d.icon} alt="" />
                                </div> */}
                                <h2>{d.title}</h2>
                                <p>{d.desc}</p>
                                <div className="iconContainer">
                                <a href={d.livesite} target="_blank"> <LanguageOutlinedIcon className="websiteIcon" /> </a>
                                <a href={d.repo} target="_blank"> <GitHubIcon className="gitHubIcon" /> </a>
                                </div>
                                
                                {/* {d.techUsed.length && (
                                    <ul className="techUsed">
                                        {d.techUsed.map((tech, i) => (
                                    <li key={i}>{tech}</li>
                                    ))}
                                    </ul>
                                    )} */}
                                    
                            </div>
                        </div>
                        <div className="right">
                            <img src={process.env.PUBLIC_URL + "/" +  d.img} alt="" />
                            <video className="vid" src={process.env.PUBLIC_URL + "/" +  d.vid} height="300" width="400" controls autoPlay={true}></video>
                        </div>
                    </div>
                    
                </div>
                ))}
            </div>
            <ArrowBackIosIcon className="arrow left" onClick={()=>handleClick("left")}/>
            <ArrowForwardIosIcon className="arrow right"  onClick={()=>handleClick("right")}/>

        
        </div>
    )
}