Next
    10.0.0.200
    Bonded RR
    Kosh
    Local 240 SSD
        \local\cache
        \docker
        \local\local
        \local\next0    320
        \local\next1    1000
        \local\next2    1000
    docker  master
        portainer 
        tdarr
        handbrake

Cube
    10.0.0.201
    Bonded RR
    Kosh
    Local 500
        \docker
        \local\local
        \local\cube0    1000
        \local\cube1    2000
        \local\cube2    3000
    docker
        portainer
        tdarr client
        Plex
        omni
        nginx
        sonarr
        radarr
        lidarr


sudo apt-get install samba -y
        sudo apt-get install cifs-utils -y

Docker token dckr_pat_5DJ9lAr5RAGbj15GbPuQHkLMez4
Plex-Token=mYvhkbcizBj-5vHWqcHo

docker run -d -p 8000:8000 -p 9443:9443 --name=portainer --restart=always -v /var/run/docker.sock:/var/run/docker.sock -v portainer_data:/data portainer/portainer-ee:latest